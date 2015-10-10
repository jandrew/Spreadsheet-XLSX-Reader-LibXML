#########1 Test File for Spreadsheet::XLSX::Reader::XMLReader::WorksheetToRow   8#########9
#!/usr/bin/env perl
my ( $lib, $test_file, $styles_file );
BEGIN{
	$ENV{PERL_TYPE_TINY_XS} = 0;
	my	$start_deeper = 1;
	$lib		= 'lib';
	$test_file	= 't/test_files/xl/';
	for my $next ( <*> ){
		if( ($next eq 't') and -d $next ){
			$start_deeper = 0;
			last;
		}
	}
	if( $start_deeper ){
		$lib		= '../../../../../../' . $lib;
		$test_file	= '../../../../../test_files/xl/';
	}
}
$| = 1;

use	Test::Most tests => 1200;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use Types::Standard qw( Bool HasMethods );
use	lib
		'../../../../../../../Log-Shiras/lib',
		$lib,
	;
use	Data::Dumper;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(
#~ ###LogSD						name_space_bounds =>{
#~ ###LogSD							UNBLOCK =>{
#~ ###LogSD								log_file => 'warn',
#~ ###LogSD							},
#~ ###LogSD							main =>{
#~ ###LogSD								UNBLOCK =>{
#~ ###LogSD									log_file => 'info',
#~ ###LogSD								},
#~ ###LogSD							},
#~ ###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow;
use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings;
use	DateTimeX::Format::Excel;
use	DateTime::Format::Flexible;
use	Type::Coercion;
use	Type::Tiny;

	$test_file	= ( @ARGV ) ? $ARGV[0] : $test_file;
my	$shared_strings_file = $test_file . 'sharedStrings.xml';
	$test_file .= 'worksheets/sheet3.xml';
	
###LogSD	my	$log_space	= 'Test';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $error_instance, $workbook_instance, $file_handle, $format_instance, $shared_strings_instance
	);
my 			@class_attributes = qw(
				file
				error_inst
				is_hidden
				workbook_instance
			);
my  		@instance_methods = qw(
				is_sheet_hidden
				is_empty_the_end
				get_group_return_type
				_set_min_col
				_min_col
				has_min_col
				_set_min_row
				_min_row
				has_min_row
				_set_max_col
				_max_col
				has_max_col
				_set_max_row
				_max_row
				has_max_row
				_set_merge_map
				_get_merge_map
				_get_row_merge_map
				get_file
				set_file
				has_file
				clear_file
				_get_row_all
				_get_next_value_cell
				_get_col_row
				is_sheet_hidden
				_is_column_hidden
				_get_row_hidden
			);
my			$answer_ref = [
				[
					{ r => 'A2', cell_row => 2, cell_col => 1, cell_xml_value => 'Hello', cell_type => 'Text' },
					{ r => 'D2', cell_row => 2, cell_col => 4, cell_xml_value => 'my', cell_type => 'Text', cell_hidden => 'column', },#
					{ r => 'C4', cell_row => 4, cell_col => 3, cell_xml_value => 'World', cell_type => 'Text', s => '7', cell_hidden => 'column', },
					{ r => 'A6', cell_row => 6, cell_col => 1, cell_xml_value => 'Hello World', cell_type => 'Text', s => '11', cell_merge => 'A6:B6' },
					{ r => 'B6', cell_row => 6, cell_col => 2, cell_type => 'Text', s => '11', cell_merge => 'A6:B6', },
					{ r => 'B7', cell_row => 7, cell_col => 2, cell_xml_value => '69', cell_type => 'Numeric', cell_hidden => 'row', },
					{ r => 'B8', cell_row => 8, cell_col => 2, cell_xml_value => '27', cell_type => 'Numeric', cell_hidden => 'row', },
					{ r => 'E8', cell_row => 8, cell_col => 5, cell_xml_value => '37145', cell_type => 'Numeric', s => 2, cell_hidden => 'row', },
					{ r => 'B9', cell_row => 9, cell_col => 2, cell_xml_value => '42', cell_type => 'Numeric', cell_formula => 'B7-B8', cell_hidden => 'row', },
					{ r => 'D10', cell_row => 10, cell_col => 4, cell_type => 'Text', cell_xml_value => ' ', s => 1, cell_hidden => 'column', },
					{ r => 'E10', cell_row => 10, cell_col => 5, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 6, cell_hidden => 'row', },
					{ r => 'F10', cell_row => 10, cell_col => 6, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 2, cell_hidden => 'row', },
					{ r => 'A11', cell_row => 11, cell_col => 1, cell_type => 'Numeric', cell_xml_value => '2.1345678901', s => 8, cell_hidden => 'row', },
					{ r => 'B12', cell_row => 12, cell_col => 2, cell_type => 'Text', cell_formula => 'IF(B11>0,"Hello","")', },
					{ r => 'D12', cell_row => 12, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'DATEVALUE(E10)', s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
					{ r => 'E12', cell_row => 12, cell_col => 5, cell_type => 'Text', s => 10, cell_merge => 'D12:E12', },
					{ r => 'C14', cell_row => 14, cell_col => 3, cell_type => 'Text', cell_xml_value => ' ', cell_hidden => 'column', },
					{ r => 'D14', cell_row => 14, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D12', s => 9, cell_hidden => 'column', },
					{ r => 'E14', cell_row => 14, cell_col => 5, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D14', s => 2, },
					'EOF',
				],
				[
					[ undef, undef, undef, undef, undef, undef,'EOR'],
					[
						{ r => 'A2', cell_row => 2, cell_col => 1, cell_xml_value => 'Hello', cell_type => 'Text' },
						undef, undef,
						{ r => 'D2', cell_row => 2, cell_col => 4, cell_xml_value => 'my', cell_type => 'Text', cell_hidden => 'column', },
						undef, undef,'EOR'
					],
					[undef, undef, undef, undef, undef, undef,'EOR'],
					[
						undef, undef,
						{ r => 'C4', cell_row => 4, cell_col => 3, cell_xml_value => 'World', cell_type => 'Text', s => '7', cell_hidden => 'column', },
						undef, undef, undef,'EOR'
					],
					[undef, undef, undef, undef, undef, undef,'EOR'],
					[
						{ r => 'A6', cell_row => 6, cell_col => 1, cell_xml_value => 'Hello World', cell_type => 'Text', s => '11', cell_merge => 'A6:B6' },
						{ r => 'B6', cell_row => 6, cell_col => 2, cell_type => 'Text', s => '11', cell_merge => 'A6:B6', },
						undef, undef, undef, undef,'EOR'
					],
					[
						undef,
						{ r => 'B7', cell_row => 7, cell_col => 2, cell_xml_value => '69', cell_type => 'Numeric', cell_hidden => 'row', },
						undef, undef, undef, undef,'EOR'
					],
					[
						undef,
						{ r => 'B8', cell_row => 8, cell_col => 2, cell_xml_value => '27', cell_type => 'Numeric', cell_hidden => 'row', },
						undef, undef,
						{ r => 'E8', cell_row => 8, cell_col => 5, cell_xml_value => '37145', cell_type => 'Numeric', s => 2, cell_hidden => 'row', },
						undef,'EOR'
					],
					[
						undef,
						{ r => 'B9', cell_row => 9, cell_col => 2, cell_xml_value => '42', cell_type => 'Numeric', cell_formula => 'B7-B8', cell_hidden => 'row', },
						undef, undef, undef, undef,'EOR'
					],
					[
						undef, undef, undef,
						{ r => 'D10', cell_row => 10, cell_col => 4, cell_type => 'Text', cell_xml_value => ' ', s => 1, cell_hidden => 'column', },
						{ r => 'E10', cell_row => 10, cell_col => 5, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 6, cell_hidden => 'row', },
						{ r => 'F10', cell_row => 10, cell_col => 6, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 2, cell_hidden => 'row', },
						'EOR'
					],
					[
						{ r => 'A11', cell_row => 11, cell_col => 1, cell_type => 'Numeric', cell_xml_value => '2.1345678901', s => 8, cell_hidden => 'row', },
						undef, undef, undef, undef, undef,'EOR'
					],
					[
						undef,
						{ r => 'B12', cell_row => 12, cell_col => 2, cell_type => 'Text', cell_formula => 'IF(B11>0,"Hello","")', },
						undef,
						{ r => 'D12', cell_row => 12, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'DATEVALUE(E10)', s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
						{ r => 'E12', cell_row => 12, cell_col => 5, cell_type => 'Text', s => 10, cell_merge => 'D12:E12', },
						undef,'EOR'
					],
					[undef, undef, undef, undef, undef, undef,'EOR'],
					[
						undef, undef,
						{ r => 'C14', cell_row => 14, cell_col => 3, cell_type => 'Text', cell_xml_value => ' ', cell_hidden => 'column', },
						{ r => 'D14', cell_row => 14, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D12', s => 9, cell_hidden => 'column', },
						{ r => 'E14', cell_row => 14, cell_col => 5, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D14', s => 2, },
						undef,'EOF'
					],
					'EOF'
				],
				[
					[],
					[
						{ r => 'A2', cell_row => 2, cell_col => 1, cell_xml_value => 'Hello', cell_type => 'Text' },undef, undef,
						{ r => 'D2', cell_row => 2, cell_col => 4, cell_xml_value => 'my', cell_type => 'Text', cell_hidden => 'column', }
					],
					[],
					[
						undef, undef,{ r => 'C4', cell_row => 4, cell_col => 3, cell_xml_value => 'World', cell_type => 'Text', s => '7', cell_hidden => 'column', }
					],
					[],
					[
						{ r => 'A6', cell_row => 6, cell_col => 1, cell_xml_value => 'Hello World', cell_type => 'Text', s => '11', cell_merge => 'A6:B6' },
						{ r => 'B6', cell_row => 6, cell_col => 2, cell_type => 'Text', s => '11', cell_merge => 'A6:B6', }
					],
					[
						undef,{ r => 'B7', cell_row => 7, cell_col => 2, cell_xml_value => '69', cell_type => 'Numeric', cell_hidden => 'row', }
					],
					[
						undef,{ r => 'B8', cell_row => 8, cell_col => 2, cell_xml_value => '27', cell_type => 'Numeric', cell_hidden => 'row', },undef, undef,
						{ r => 'E8', cell_row => 8, cell_col => 5, cell_xml_value => '37145', cell_type => 'Numeric', s => 2, cell_hidden => 'row', }
					],
					[
						undef,{ r => 'B9', cell_row => 9, cell_col => 2, cell_xml_value => '42', cell_type => 'Numeric', cell_formula => 'B7-B8', cell_hidden => 'row', }
					],
					[
						undef, undef, undef,{ r => 'D10', cell_row => 10, cell_col => 4, cell_type => 'Text', cell_xml_value => ' ', s => 1, cell_hidden => 'column', },
						{ r => 'E10', cell_row => 10, cell_col => 5, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 6, cell_hidden => 'row', },
						{ r => 'F10', cell_row => 10, cell_col => 6, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 2, cell_hidden => 'row', },
					],
					[
						{ r => 'A11', cell_row => 11, cell_col => 1, cell_type => 'Numeric', cell_xml_value => '2.1345678901', s => 8, cell_hidden => 'row', }
					],
					[
						undef,
						{ r => 'B12', cell_row => 12, cell_col => 2, cell_type => 'Text', cell_formula => 'IF(B11>0,"Hello","")', }, undef,
						{ r => 'D12', cell_row => 12, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'DATEVALUE(E10)', s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
						{ r => 'E12', cell_row => 12, cell_col => 5, cell_type => 'Text', s => 10, cell_merge => 'D12:E12', }
					],
					[],
					[
						undef, undef,{ r => 'C14', cell_row => 14, cell_col => 3, cell_type => 'Text', cell_xml_value => ' ', cell_hidden => 'column', },
						{ r => 'D14', cell_row => 14, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D12', s => 9, cell_hidden => 'column', },
						{ r => 'E14', cell_row => 14, cell_col => 5, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D14', s => 2, }
					],	
					'EOF',
				],
				[
					{ r => 'A2', cell_row => 2, cell_col => 1, cell_xml_value => 'Hello', cell_type => 'Text' },
					undef, undef,
					{ r => 'D2', cell_row => 2, cell_col => 4, cell_xml_value => 'my', cell_type => 'Text', cell_hidden => 'column', },
					undef,
					undef, undef,
					{ r => 'C4', cell_row => 4, cell_col => 3, cell_xml_value => 'World', cell_type => 'Text', s => '7', cell_hidden => 'column', },
					undef,
					{ r => 'A6', cell_row => 6, cell_col => 1, cell_xml_value => 'Hello World', cell_type => 'Text', s => '11', cell_merge => 'A6:B6' },
					{ r => 'B6', cell_row => 6, cell_col => 2, cell_type => 'Text', s => '11', cell_merge => 'A6:B6', },
					undef,
					{ r => 'B7', cell_row => 7, cell_col => 2, cell_xml_value => '69', cell_type => 'Numeric', cell_hidden => 'row', },
					undef,
					{ r => 'B8', cell_row => 8, cell_col => 2, cell_xml_value => '27', cell_type => 'Numeric', cell_hidden => 'row', },
					undef, undef,
					{ r => 'E8', cell_row => 8, cell_col => 5, cell_xml_value => '37145', cell_type => 'Numeric', s => 2, cell_hidden => 'row', },
					undef,
					{ r => 'B9', cell_row => 9, cell_col => 2, cell_xml_value => '42', cell_type => 'Numeric', cell_formula => 'B7-B8', cell_hidden => 'row', },
					undef, undef, undef,
					{ r => 'D10', cell_row => 10, cell_col => 4, cell_type => 'Text', cell_xml_value => ' ', s => 1, cell_hidden => 'column', },
					{ r => 'E10', cell_row => 10, cell_col => 5, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 6, cell_hidden => 'row', },
					{ r => 'F10', cell_row => 10, cell_col => 6, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 2, cell_hidden => 'row', },
					{ r => 'A11', cell_row => 11, cell_col => 1, cell_type => 'Numeric', cell_xml_value => '2.1345678901', s => 8, cell_hidden => 'row', },
					undef,
					{ r => 'B12', cell_row => 12, cell_col => 2, cell_type => 'Text', cell_formula => 'IF(B11>0,"Hello","")', }, undef,
					{ r => 'D12', cell_row => 12, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'DATEVALUE(E10)', s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
					{ r => 'E12', cell_row => 12, cell_col => 5, cell_type => 'Text', s => 10, cell_merge => 'D12:E12', },
					undef,
					undef, undef,
					{ r => 'C14', cell_row => 14, cell_col => 3, cell_type => 'Text', cell_xml_value => ' ', cell_hidden => 'column', },
					{ r => 'D14', cell_row => 14, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D12', s => 9, cell_hidden => 'column', },
					{ r => 'E14', cell_row => 14, cell_col => 5, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D14', s => 2, },
					'EOF',
				],
				[
					'EOR',
					{ r => 'A2', cell_row => 2, cell_col => 1, cell_xml_value => 'Hello', cell_type => 'Text' },
					undef, undef,
					{ r => 'D2', cell_row => 2, cell_col => 4, cell_xml_value => 'my', cell_type => 'Text', cell_hidden => 'column', },
					,'EOR',
					'EOR',
					undef, undef,
					{ r => 'C4', cell_row => 4, cell_col => 3, cell_xml_value => 'World', cell_type => 'Text', s => '7', cell_hidden => 'column', },
					'EOR',
					'EOR',
					{ r => 'A6', cell_row => 6, cell_col => 1, cell_xml_value => 'Hello World', cell_type => 'Text', s => '11', cell_merge => 'A6:B6' },
					{ r => 'B6', cell_row => 6, cell_col => 2, cell_type => 'Text', s => '11', cell_merge => 'A6:B6', },
					'EOR',
					undef,
					{ r => 'B7', cell_row => 7, cell_col => 2, cell_xml_value => '69', cell_type => 'Numeric', cell_hidden => 'row', },
					'EOR',
					undef,
					{ r => 'B8', cell_row => 8, cell_col => 2, cell_xml_value => '27', cell_type => 'Numeric', cell_hidden => 'row', },
					undef, undef,
					{ r => 'E8', cell_row => 8, cell_col => 5, cell_xml_value => '37145', cell_type => 'Numeric', s => 2, cell_hidden => 'row', },
					'EOR',
					undef,
					{ r => 'B9', cell_row => 9, cell_col => 2, cell_xml_value => '42', cell_type => 'Numeric', cell_formula => 'B7-B8', cell_hidden => 'row', },
					'EOR',
					undef, undef, undef,
					{ r => 'D10', cell_row => 10, cell_col => 4, cell_type => 'Text', cell_xml_value => ' ', s => 1, cell_hidden => 'column', },
					{ r => 'E10', cell_row => 10, cell_col => 5, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 6, cell_hidden => 'row', },
					{ r => 'F10', cell_row => 10, cell_col => 6, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 2, cell_hidden => 'row', },
					'EOR',
					{ r => 'A11', cell_row => 11, cell_col => 1, cell_type => 'Numeric', cell_xml_value => '2.1345678901', s => 8, cell_hidden => 'row', },
					'EOR',
					undef,
					{ r => 'B12', cell_row => 12, cell_col => 2, cell_type => 'Text', cell_formula => 'IF(B11>0,"Hello","")', }, undef,
					{ r => 'D12', cell_row => 12, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'DATEVALUE(E10)', s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
					{ r => 'E12', cell_row => 12, cell_col => 5, cell_type => 'Text', s => 10, cell_merge => 'D12:E12', },
					'EOR',
					'EOR',
					undef, undef,
					{ r => 'C14', cell_row => 14, cell_col => 3, cell_type => 'Text', cell_xml_value => ' ', cell_hidden => 'column', },
					{ r => 'D14', cell_row => 14, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D12', s => 9, cell_hidden => 'column', },
					{ r => 'E14', cell_row => 14, cell_col => 5, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D14', s => 2, },
					'EOF',
				],
				[
					[],
					[
						{ r => 'A2', cell_row => 2, cell_col => 1, cell_xml_value => 'Hello', cell_type => 'Text' },undef, undef,
						{ r => 'D2', cell_row => 2, cell_col => 4, cell_xml_value => 'my', cell_type => 'Text', cell_hidden => 'column', },
					],
					[],
					[
						undef, undef,{ r => 'C4', cell_row => 4, cell_col => 3, cell_xml_value => 'World', cell_type => 'Text', s => '7', cell_hidden => 'column', },
					],
					[],
					[
						{ r => 'A6', cell_row => 6, cell_col => 1, cell_xml_value => 'Hello World', cell_type => 'Text', s => '11', cell_merge => 'A6:B6' },
						{ r => 'B6', cell_row => 6, cell_col => 2, cell_type => 'Text', s => '11', cell_merge => 'A6:B6', },
					],
					[
						undef,{ r => 'B7', cell_row => 7, cell_col => 2, cell_xml_value => '69', cell_type => 'Numeric', cell_hidden => 'row', },
					],
					[
						undef,{ r => 'B8', cell_row => 8, cell_col => 2, cell_xml_value => '27', cell_type => 'Numeric', cell_hidden => 'row', },undef, undef,
						{ r => 'E8', cell_row => 8, cell_col => 5, cell_xml_value => '37145', cell_type => 'Numeric', s => 2, cell_hidden => 'row', },
					],
					[
						undef,{ r => 'B9', cell_row => 9, cell_col => 2, cell_xml_value => '42', cell_type => 'Numeric', cell_formula => 'B7-B8', cell_hidden => 'row', },
					],
					[
						undef, undef, undef,{ r => 'D10', cell_row => 10, cell_col => 4, cell_type => 'Text', cell_xml_value => ' ', s => 1, cell_hidden => 'column', },
						{ r => 'E10', cell_row => 10, cell_col => 5, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 6, cell_hidden => 'row', },
						{ r => 'F10', cell_row => 10, cell_col => 6, cell_type => 'Text', cell_xml_value => '2/6/2011', s => 2, cell_hidden => 'row', },
					],
					[
						{ r => 'A11', cell_row => 11, cell_col => 1, cell_type => 'Numeric', cell_xml_value => '2.1345678901', s => 8, cell_hidden => 'row', },
					],
					[
						undef,
						{ r => 'B12', cell_row => 12, cell_col => 2, cell_type => 'Text', cell_formula => 'IF(B11>0,"Hello","")', }, undef,
						{ r => 'D12', cell_row => 12, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'DATEVALUE(E10)', s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
						{ r => 'E12', cell_row => 12, cell_col => 5, cell_type => 'Text', s => 10, cell_merge => 'D12:E12', },
					],
					[],
					[
						undef, undef,{ r => 'C14', cell_row => 14, cell_col => 3, cell_type => 'Text', cell_xml_value => ' ', cell_hidden => 'column', },
						{ r => 'D14', cell_row => 14, cell_col => 4, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D12', s => 9, cell_hidden => 'column', },
						{ r => 'E14', cell_row => 14, cell_col => 5, cell_type => 'Numeric', cell_xml_value => '39118', cell_formula => 'D14', s => 2, },
					],	
					'EOF',
				],
				[ 0, 0, 1, 1, 0, 0 ],
				[ undef, undef, 0, undef, 0, undef, 0, 1, 1, 1, 1, 1, 0, undef, 0, undef ],
			];
###LogSD	$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{
has_attribute_ok
			'Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow', $_,
										"Check that Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow has the -$_- attribute"
} 			@class_attributes;

lives_ok{
			$error_instance = Spreadsheet::XLSX::Reader::LibXML::Error->new( should_warn => 0 );
			$format_instance = Spreadsheet::XLSX::Reader::LibXML::FmtDefault->new(
										epoch_year	=> 1904,
										error_inst	=> $error_instance,
				###LogSD				log_space	=> 'Test',
									);
			$shared_strings_instance =	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings->new(
										group_return_type	=> 'xml_value',
										file			=> $shared_strings_file,
										error_inst 		=> Spreadsheet::XLSX::Reader::LibXML::Error->new(
											#~ should_warn		=> 1,
											should_warn		=> 0,# to turn off cluck when the error is set
										),
				###LogSD				log_space	=> 'Test',
									);
			$workbook_instance	= build_instance(
									package		=> 'WorkbookInstance',
									add_methods =>{
										counting_from_zero			=> sub{ return 0 },
										boundary_flag_setting		=> sub{},
										change_boundary_flag		=> sub{},
										_has_shared_strings_file	=> sub{ return 1 },
										_has_styles_file			=> sub{},
										get_format_position			=> sub{},
										get_epoch_year				=> sub{ return 1904 },
										get_group_return_type		=> sub{},
										set_group_return_type		=> sub{},
										get_date_behavior			=> sub{},
										set_date_behavior			=> sub{},
										get_empty_return_type		=> sub{ return 'undef_string' },
										get_values_only				=> sub{},
										set_values_only				=> sub{},
									},
									add_attributes =>{
										error_inst =>{
											isa			=> 	HasMethods[qw(
																error set_error clear_error set_warnings if_warn
															) ],
											clearer		=> '_clear_error_inst',
											reader		=> 'get_error_inst',
											required	=> 1,
											handles =>[ qw(
												error set_error clear_error set_warnings if_warn
											) ],
										},
										empty_is_end =>{
											isa		=> Bool,
											writer	=> 'set_empty_is_end',
											reader	=> 'is_empty_the_end',
											default	=> 0,
										},
										from_the_edge =>{
											isa		=> Bool,
											reader	=> '_starts_at_the_edge',
											writer	=> 'set_from_the_edge',
											default	=> 1,
										},
										format_instance =>{
											isa		=> HasMethods[qw( 
															set_error_inst				set_excel_region
															set_target_encoding			get_defined_excel_format
															set_defined_excel_formats	change_output_encoding
															set_epoch_year				set_cache_behavior
															set_date_behavior			get_defined_conversion		
															parse_excel_format_string							)],	
											writer	=> 'set_format_instance',
											reader	=> 'get_format_instance',
											handles =>[qw(
															get_defined_excel_format 	parse_excel_format_string
															change_output_encoding		)],
										},
										_shared_strings_instance =>{
											isa			=> HasMethods[ 'get_shared_string_position' ],
											predicate	=> '_has_shared_strings_file',
											writer		=> '_set_shared_strings_instance',
											reader		=> '_get_shared_strings_instance',
											clearer		=> '_clear_shared_strings',
											handles		=>{
												'get_shared_string_position' => 'get_shared_string_position',
												_demolish_shared_strings => 'DEMOLISH',
											},
										},
									},
									error_inst => $error_instance,
									format_instance => $format_instance,
									_shared_strings_instance => $shared_strings_instance,
								);
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow->new(
				file				=> $test_file,
				error_inst			=> $error_instance,
				workbook_instance	=> $workbook_instance,
				is_hidden 			=> 0,
			###LogSD	log_space	=> 'Test',
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a new WorksheetToRow instance";

map{
can_ok		$test_instance, $_,
} 			@instance_methods;
is			$test_instance->_min_row, 1,
										"check that it knows what the lowest row number is";
is			$test_instance->_min_col, 1,
										"check that it knows what the lowest column number is";
is			$test_instance->_max_row, undef,
										"check that it knows what the highest row number is (not)";
is			$test_instance->_max_col, undef,
										"check that it knows what the highest column number is (not)";
										
explain									"read through value cells ...";
			my $test = 0;
			for my $y (1..2){
			my $result;
explain									"Running cycle: $y";
			my $x = 0;
			while( !$result or $result ne 'EOF' ){
				
###LogSD	my $expose = 20; my $iteration = 1;
###LogSD	if( $x == $expose and $y == $iteration ){
###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
#~ ###LogSD				_get_next_value_cell =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'trace',
###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $x > ($expose + 0) and $y == $iteration ){
###LogSD		exit 1;
###LogSD	}

lives_ok{	$result = $test_instance->_get_next_value_cell }
										"_get_next_value_cell test -$test- iteration -$y- from sheet position: $x";
			#~ print Dumper( $result );
###LogSD	$phone->talk( level => 'debug', message => [ "result at position -$x- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$test]->[$x], ] );
is_deeply	$result, $answer_ref->[$test]->[$x],"..........and see if test -$test- iteration -$y- from sheet position -$x- has good info";
#~ explain									Dumper( $test_instance->_get_all_positions );
			$x++;
#~ explain									"Checking next x: $x";
			}
#~ explain									"Checking y after: $y";
			}
			$test++;
explain									"Finished value cell reading";
explain									"read row columns through cells in sequence...";
			for my $y (1..3){
explain									"Running cycle: $y";
			if( $y == 3 ){
lives_ok{
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow->new(
				file				=> $test_file,
				error_inst			=> $error_instance,
				workbook_instance	=> $workbook_instance,
				is_hidden 			=> 0,
			###LogSD	log_space	=> 'Test',
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a new WorksheetToRow instance";
			}
			my $y_dim = 1;
			#~ my $x = 0;
			my	$result = undef;
			while( !$result or $result ne 'EOF' ){
			my	$x_dim = 1;
				$result = undef;
			while( !$result or ($result ne 'EOR' and $result ne 'EOF') ){
				
###LogSD	my $expose_x = 7;
###LogSD	my $expose_y = 16;
###LogSD	my $expose_dim = 1;
###LogSD	if( $x_dim == $expose_x and $y_dim == $expose_y and $y == $expose_dim ){
###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
#~ ###LogSD				WorksheetToRow =>{
#~ ###LogSD					_go_to_or_past_row =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
###LogSD		} );
###LogSD	}


lives_ok{	$result = $test_instance->_get_col_row( $x_dim, $y_dim  ) }
										"_get_col_row data for test -$test- and iteration -$y- at column -$x_dim- and row -$y_dim-";
###LogSD	$phone->talk( level => 'trace', message => [ "result for column -$y_dim- and row -$x_dim- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$test]->[$y_dim-1]->[$x_dim-1], ] );
			if( $y == 3 and $result and $result eq 'EOR' ){
pass									"...........and see if test -$test- and iteration -$y- at column -$x_dim- and row -$y_dim- found EOR";
			}else{
is_deeply	$result, $answer_ref->[$test]->[$y_dim-1]->[$x_dim-1],
										"...........and see if test -$test- and iteration -$y- at column -$x_dim- and row -$y_dim- returns good info";#: " . Dumper( $answer_ref->[$test]->[$y_dim-1]->[$x_dim-1] );
			}
###LogSD	if( $x_dim > $expose_x and $y_dim == $expose_y and $y == $expose_dim ){
###LogSD		exit 1;
###LogSD	}
			$x_dim++;
			}
			$y_dim++;
			}
			}
			$test++;
explain									"read rows through sheet in sequence...";
			for my $y (1..3){
explain									"Running cycle: $y";
			if( $y == 3 ){
lives_ok{
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow->new(
				file				=> $test_file,
				error_inst			=> $error_instance,
				workbook_instance	=> $workbook_instance,
				is_hidden 			=> 0,
			###LogSD	log_space	=> 'Test',
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a new WorksheetToRow instance";
			}
			my $result = undef;
			my $y_dim = 1;
			my $x = 0;
			while( !$result or $result ne 'EOF' ){
				
###LogSD	my $expose_y = 17;
###LogSD	my $iteration = 1;
###LogSD	if( $y_dim == $expose_y and $y == $iteration ){
###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'trace',
###LogSD					},
#~ ###LogSD				parse_element =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'warn',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD				XMLReader =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'warn',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $y_dim > $expose_y and $y == $iteration ){
###LogSD		exit 1;
###LogSD	}

lives_ok{	$result = $test_instance->_get_row_all( $y_dim ) }
										"For test -$test- and iteration -$y- collecting _get_row_all data for row -$y_dim-";
###LogSD	$phone->talk( level => 'trace', message => [ "result for row -$y_dim- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$test]->[$x], ] );
is_deeply	$result, $answer_ref->[$test]->[$x++],"..and see if test -$test- and iteration -$y- has good info for row -$y_dim-";
			$y_dim++;
			}
			}

lives_ok{
			$workbook_instance->set_empty_is_end( 1 );
			$workbook_instance->set_from_the_edge( 0 );
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow->new(
				file				=> $test_file,
				error_inst			=> $error_instance,
				workbook_instance	=> $workbook_instance,
				is_hidden			=> 0,
			###LogSD	log_space	=> 'Test',
			);
###LogSD	$phone->talk( level => 'trace', message =>[ "Loaded new test instance - without the edges" ] );
}										"Build a Worksheet instance with the edges cut off";
			$test++;
#~ explain									"read through cells without edges in sequence...";
			#~ for my $y (1..2){
			#~ my $result = undef;
#~ explain									"Running cycle: $y";
			#~ my $x = 218;
			#~ while( $x < 257 and (!$result or $result ne 'EOF') ){
			#~ my	$position = $x - 218;
#~ lives_ok{	$result = $test_instance->_get_next_cell }
										#~ "Collecting data from sheet position: $position";
#~ ###LogSD	$phone->talk( level => 'trace', message => [ "result at position -$position- is:", $result,
#~ ###LogSD		'Against answer:', $answer_ref->[$x], ] );
#~ is_deeply	$result, $answer_ref->[$x++],"..and see if it has good info";
			#~ }
			#~ }
			$test++;
explain									"read row columns through cells without edges in sequence...";
			for my $y (1..3){
explain									"Running cycle: $y";
			if( $y == 3 ){
lives_ok{
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow->new(
				file				=> $test_file,
				error_inst			=> $error_instance,
				workbook_instance	=> $workbook_instance,
				is_hidden 			=> 0,
			###LogSD	log_space	=> 'Test',
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a new WorksheetToRow instance";
			}
			my $y_dim = 1;
			my $x = 0;
			my	$result = undef;
			while( !$result or $result ne 'EOF' ){
			my	$x_dim = 1;
				$result = undef;
			while( !$result or ($result ne 'EOR' and $result ne 'EOF') ){
				
###LogSD	my $expose_x = 5;
###LogSD	my $expose_y = 2;
###LogSD	if( $x_dim == $expose_x and $y_dim == $expose_y ){
###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'trace',
###LogSD					},
#~ ###LogSD				parse_element =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'warn',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD				XMLReader =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'warn',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $y_dim > $expose_y or ($y_dim == $expose_y and $x_dim > $expose_x ) ){
###LogSD		exit 1;
###LogSD	}

lives_ok{	$result = $test_instance->_get_col_row( $x_dim, $y_dim  ) }
										"_get_col_row data for test -$test- and iteration -$y- at column -$x_dim- and row -$y_dim-";
###LogSD	$phone->talk( level => 'trace', message => [ "result for column -$y_dim- and row -$x_dim- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$test]->[$x], ] );
is_deeply	$result, $answer_ref->[$test]->[$x],"...........and see if test -$test- and iteration -$y- at column -$x_dim- and row -$y_dim- returns good info";
			$x++;
			$x_dim++;
			}
			$y_dim++;
			}
			}
			$test++;
explain									"read rows through sheet without edges in sequence...";
			for my $y (1..3){
explain									"Running cycle: $y";
			if( $y == 3 ){
lives_ok{
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::WorksheetToRow->new(
				file				=> $test_file,
				error_inst			=> $error_instance,
				workbook_instance	=> $workbook_instance,
				is_hidden 			=> 0,
			###LogSD	log_space	=> 'Test',
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a new WorksheetToRow instance";
			}
			my $result = undef;
			my $y_dim = 1;
			my $x = 0;
			while( !$result or $result ne 'EOF' ){
				
###LogSD	my $expose_y = 20;
###LogSD	if( $y_dim == $expose_y ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			Test =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'trace',
###LogSD					},
###LogSD				parse_element =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				XMLReader =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD	}

###LogSD	elsif( $y_dim > $expose_y ){
###LogSD		exit 1;
###LogSD	}

lives_ok{	$result = $test_instance->_get_row_all( $y_dim ) }
										"For test -$test- and iteration -$y- collecting _get_row_all data for row -$y_dim-";
###LogSD	$phone->talk( level => 'trace', message => [ "result for row -$y_dim- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$test]->[$x], ] );
is_deeply	$result, $answer_ref->[$test]->[$x++],"..and see if test -$test- and iteration -$y- has good info for row -$y_dim-";
			$y_dim++;
			}
			}
			$test++;
is			$test_instance->is_sheet_hidden, 0,
										'Check if the sheet is hidden (Not)';
is_deeply	[ $test_instance->_is_column_hidden( 1 .. 6 ) ], $answer_ref->[$test],#[ 0, 0, 1, 1, 0, 0 ],
										'Check that the sheet knows which columns are hidden - by number';
			$test++;
#~ ###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
#~ ###LogSD				is_row_hidden =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'trace',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD		} );
			for my $row ( 0..15 ){
is			$test_instance->_get_row_hidden( $row ),  $answer_ref->[$test]->[$row],
										"For test -$test- check that the sheet knows the hidden state of row: $row";
			}
is			$test_instance->_max_row, 14,
										"check that it knows what the highest row number is: 14";
is			$test_instance->_max_col, 6,
										"check that it knows what the highest column number is: 6";
explain 								"...Test Done";
done_testing();

###LogSD	package Print::Log;
###LogSD	use Data::Dumper;
###LogSD	sub new{
###LogSD		bless {}, shift;
###LogSD	}
###LogSD	sub add_line{
###LogSD		shift;
###LogSD		my @input = ( ref $_[0]->{message} eq 'ARRAY' ) ? 
###LogSD						@{$_[0]->{message}} : $_[0]->{message};
###LogSD		my ( @print_list, @initial_list );
###LogSD		no warnings 'uninitialized';
###LogSD		for my $value ( @input ){
###LogSD			push @initial_list, (( ref $value ) ? Dumper( $value ) : $value );
###LogSD		}
###LogSD		for my $line ( @initial_list ){
###LogSD			$line =~ s/\n$//;
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "| level - %-6s | name_space - %-s\n| line  - %04d   | file_name  - %-s\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{level}, $_[0]->{name_space},
###LogSD					$_[0]->{line}, $_[0]->{filename},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;