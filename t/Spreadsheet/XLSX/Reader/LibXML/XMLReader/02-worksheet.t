#########1 Test File for Spreadsheet::XLSX::Reader::XMLReader::Worksheet        8#########9
#!/usr/bin/env perl
my ( $lib, $test_file, $styles_file );
BEGIN{
	$ENV{PERL_TYPE_TINY_XS} = 0;
	my	$start_deeper = 1;
	$lib		= 'lib';
	$test_file	= 't/test_files/xl/worksheets/';
	for my $next ( <*> ){
		if( ($next eq 't') and -d $next ){
			$start_deeper = 0;
			last;
		}
	}
	if( $start_deeper ){
		$lib		= '../../../../../../' . $lib;
		$test_file	= '../../../../../test_files/xl/worksheets/';
	}
}
$| = 1;

use	Test::Most tests => 1338;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use Types::Standard qw( Bool HasMethods );
use	lib
		'../../../../../../../Log-Shiras/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'warn',
###LogSD							},
###LogSD							main =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'info',
###LogSD								},
###LogSD							},
###LogSD							Test =>{
###LogSD								_set_file_name =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
#~ use	Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet;
use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use	DateTimeX::Format::Excel;
use	DateTime::Format::Flexible;
use	Type::Coercion;
use	Type::Tiny;

	$test_file	= ( @ARGV ) ? $ARGV[0] : $test_file;
	$test_file .= 'sheet3.xml';
	
###LogSD	my	$log_space	= 'Test';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $error_instance, $workbook_instance, $file_handle, $format_instance,
	);
my 			@class_attributes = qw(
				file
				error_inst
				sheet_rel_id
				sheet_id
				sheet_position
				sheet_name
			);
my  		@instance_methods = qw(
				rel_id
				sheet_id
				position
				get_name
				set_empty_is_end
				is_empty_the_end
				row_range
				col_range
				min_col
				has_min_col
				min_row
				has_min_row
				max_col
				has_max_col
				max_row
				has_max_row
				get_file
				set_file
				has_file
				clear_file
				start_reading
				error
				clear_error
				set_warnings
				if_warn
				parse_column_row
				build_cell_label
				parse_element
				_get_next_value_cell
				_get_next_cell
				_get_col_row
				_get_row_all
				get_merged_areas
				is_sheet_hidden
				is_row_hidden
				is_column_hidden
			);
my			$answer_ref = [
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's', cell_hidden => 'column', },#
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's', cell_hidden => 'column', },
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, cell_hidden => 'row', },
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, cell_hidden => 'row', },
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2, cell_hidden => 'row', },
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, cell_hidden => 'row', },
				{ r => 'D10', row => 10, col => 4, v =>{ raw_text => '3' }, t => 's', s => 1, cell_hidden => 'column', },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, cell_hidden => 'row', },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', cell_hidden => 'row', },
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, cell_hidden => 'row', },
				{ r => 'B12', row => 12, col => 2, v =>{ raw_text => undef }, f =>{ raw_text => 'IF(B11>0,"Hello","")' }, },
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', cell_hidden => 'column', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, cell_hidden => 'column', },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, },
				'EOF',
				undef, undef, undef, undef, undef, undef,
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				undef, undef,
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's', cell_hidden => 'column', },
				undef, undef,
				undef, undef, undef, undef, undef, undef,
				undef, undef,
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's', cell_hidden => 'column', },
				undef, undef, undef,
				undef, undef, undef, undef, undef, undef,
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				undef, undef, undef, undef,
				undef,
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, cell_hidden => 'row', },
				undef, undef, undef, undef,
				undef,
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, cell_hidden => 'row', },
				undef, undef,
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2, cell_hidden => 'row', },
				undef,
				undef,
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, cell_hidden => 'row', },
				undef, undef, undef, undef,
				undef, undef, undef,
				{ r => 'D10', row => 10, col => 4, v =>{ raw_text => '3' }, t => 's', s => 1, cell_hidden => 'column', },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, cell_hidden => 'row', },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', cell_hidden => 'row', },
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, cell_hidden => 'row', },
				undef, undef, undef, undef, undef,
				undef,
				{ r => 'B12', row => 12, col => 2, v =>{ raw_text => undef }, f =>{ raw_text => 'IF(B11>0,"Hello","")' }, },
				undef,
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				undef,
				undef, undef, undef, undef, undef, undef,
				undef, undef,
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', cell_hidden => 'column', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, cell_hidden => 'column', },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, },
				undef,
				'EOF',
				undef, undef, undef, undef, undef, undef,'EOR',
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				undef, undef,
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's', cell_hidden => 'column', },
				undef, undef,'EOR',
				undef, undef, undef, undef, undef, undef,'EOR',
				undef, undef,
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's', cell_hidden => 'column', },
				undef, undef, undef,'EOR',
				undef, undef, undef, undef, undef, undef,'EOR',
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				undef, undef, undef, undef,'EOR',
				undef,
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, cell_hidden => 'row', },
				undef, undef, undef, undef,'EOR',
				undef,
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, cell_hidden => 'row', },
				undef, undef,
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2, cell_hidden => 'row', },
				undef,'EOR',
				undef,
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, cell_hidden => 'row', },
				undef, undef, undef, undef,'EOR',
				undef, undef, undef,
				{ r => 'D10', row => 10, col => 4, v =>{ raw_text => '3' }, t => 's', s => 1, cell_hidden => 'column', },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, cell_hidden => 'row', },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', cell_hidden => 'row', },
				'EOR',
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, cell_hidden => 'row', },
				undef, undef, undef, undef, undef,'EOR',
				undef,
				{ r => 'B12', row => 12, col => 2, v =>{ raw_text => undef }, f =>{ raw_text => 'IF(B11>0,"Hello","")' }, },
				undef,
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				undef,'EOR',
				undef, undef, undef, undef, undef, undef,'EOR',
				undef, undef,
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', cell_hidden => 'column', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, cell_hidden => 'column', },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, },
				undef,'EOF',
				[undef, undef, undef, undef, undef, undef,],
				[
					{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },undef, undef,
					{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's', cell_hidden => 'column', },undef, undef,
				],
				[undef, undef, undef, undef, undef, undef,],
				[
					undef, undef,{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's', cell_hidden => 'column', }, undef, undef, undef,
				],
				[undef, undef, undef, undef, undef, undef,],
				[
					{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
					{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', }, undef, undef, undef, undef,
				],
				[
					undef,{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, cell_hidden => 'row', }, undef, undef, undef, undef,
				],
				[
					undef,{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, cell_hidden => 'row', },undef, undef,
					{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2, cell_hidden => 'row', }, undef,
				],
				[
					undef,{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, cell_hidden => 'row', }, undef, undef, undef, undef,
				],
				[
					undef, undef, undef,{ r => 'D10', row => 10, col => 4, v =>{ raw_text => '3' }, t => 's', s => 1, cell_hidden => 'column', },
					{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, cell_hidden => 'row', },
					{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', cell_hidden => 'row', },
				],
				[
					{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, cell_hidden => 'row', }, undef, undef, undef, undef, undef,
				],
				[
					undef,
					{ r => 'B12', row => 12, col => 2, v =>{ raw_text => undef }, f =>{ raw_text => 'IF(B11>0,"Hello","")' }, }, undef,
					{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
					{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', }, undef,
				],
				[undef, undef, undef, undef, undef, undef,],
				[
					undef, undef,{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', cell_hidden => 'column', },
					{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, cell_hidden => 'column', },
					{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, }, undef,
				],	
				'EOF',
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				undef, undef,
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's', cell_hidden => 'column', },
				undef,
				undef, undef,
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's', cell_hidden => 'column', },
				undef,
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				undef,
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, cell_hidden => 'row', },
				undef,
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, cell_hidden => 'row', },
				undef, undef,
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2, cell_hidden => 'row', },
				undef,
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, cell_hidden => 'row', },
				undef, undef, undef,
				{ r => 'D10', row => 10, col => 4, v =>{ raw_text => '3' }, t => 's', s => 1, cell_hidden => 'column', },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, cell_hidden => 'row', },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', cell_hidden => 'row', },
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, cell_hidden => 'row', },
				undef,
				{ r => 'B12', row => 12, col => 2, v =>{ raw_text => undef }, f =>{ raw_text => 'IF(B11>0,"Hello","")' }, }, undef,
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				undef,
				undef, undef,
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', cell_hidden => 'column', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, cell_hidden => 'column', },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2 },
				'EOF',
				'EOR',
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				undef, undef,
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's', cell_hidden => 'column', },
				'EOR',
				'EOR',
				undef, undef,
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's', cell_hidden => 'column', },
				'EOR',
				'EOR',
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				'EOR',
				undef,
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, cell_hidden => 'row', },
				'EOR',
				undef,
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, cell_hidden => 'row', },
				undef, undef,
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2, cell_hidden => 'row', },
				'EOR',
				undef,
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, cell_hidden => 'row', },
				'EOR',
				undef, undef, undef,
				{ r => 'D10', row => 10, col => 4, v =>{ raw_text => '3' }, t => 's', s => 1, cell_hidden => 'column', },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, cell_hidden => 'row', },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', cell_hidden => 'row', },
				'EOR',
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, cell_hidden => 'row', },
				'EOR',
				undef,
				{ r => 'B12', row => 12, col => 2, v =>{ raw_text => undef }, f =>{ raw_text => 'IF(B11>0,"Hello","")' }, }, undef,
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				'EOR',
				'EOR',
				undef, undef,
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', cell_hidden => 'column', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, cell_hidden => 'column', },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2 },
				'EOF',
				[],
				[
					{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },undef, undef,
					{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's', cell_hidden => 'column', },
				],
				[],
				[
					undef, undef,{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's', cell_hidden => 'column', },
				],
				[],
				[
					{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
					{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				],
				[
					undef,{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, cell_hidden => 'row', },
				],
				[
					undef,{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, cell_hidden => 'row', },undef, undef,
					{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2, cell_hidden => 'row', },
				],
				[
					undef,{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, cell_hidden => 'row', },
				],
				[
					undef, undef, undef,{ r => 'D10', row => 10, col => 4, v =>{ raw_text => '3' }, t => 's', s => 1, cell_hidden => 'column', },
					{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, cell_hidden => 'row', },
					{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', cell_hidden => 'row', },
				],
				[
					{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, cell_hidden => 'row', },
				],
				[
					undef,
					{ r => 'B12', row => 12, col => 2, v =>{ raw_text => undef }, f =>{ raw_text => 'IF(B11>0,"Hello","")' }, }, undef,
					{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12', cell_hidden => 'column', },
					{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				],
				[],
				[
					undef, undef,{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', cell_hidden => 'column', },
					{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, cell_hidden => 'column', },
					{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, },
				],	
				'EOF',
			];
###LogSD	$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{
has_attribute_ok
			'Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet', $_,
										"Check that Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet has the -$_- attribute"
} 			@class_attributes;

lives_ok{
			$error_instance		= 	Spreadsheet::XLSX::Reader::LibXML::Error->new( should_warn => 0 );
			$format_instance	=  	Spreadsheet::XLSX::Reader::LibXML::FmtDefault->new(
										epoch_year	=> 1904,
										error_inst	=> $error_instance,
				###LogSD				log_space	=> 'Test',
									);
			$workbook_instance	= build_instance(
									package		=> 'WorkbookInstance',
									add_methods =>{
										counting_from_zero			=> sub{ return 0 },
										boundary_flag_setting		=> sub{},
										change_boundary_flag		=> sub{},
										_has_shared_strings_file	=> sub{ return 1 },
										get_shared_string_position	=> sub{},
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
									},
									error_inst => $error_instance,
									format_instance => $format_instance,
								);
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet->new(
				file				=> $test_file,
				error_inst			=> $error_instance,
				sheet_name			=> 'Sheet3',
				workbook_instance	=> $workbook_instance,
				is_hidden 			=> 0,
			###LogSD	log_space	=> 'Test',
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a new Worksheet instance";
###LogSD		$phone->talk( level => 'debug', message => [ "Max row is:" . $test_instance->max_row ] );
map{
can_ok		$test_instance, $_,
} 			@instance_methods;
is			$test_instance->min_row, 1,
										"check that it knows what the lowest row number is";
is			$test_instance->min_col, 1,
										"check that it knows what the lowest column number is";
is			$test_instance->max_row, 14,
										"check that it knows what the highest row number is";
is			$test_instance->max_col, 6,
										"check that it knows what the highest column number is";
is_deeply	[$test_instance->row_range], [1,14],
										"check for a correct row range";
is_deeply	[$test_instance->col_range], [1,6],
										"check for a correct column range";
										
explain									"read through value cells ...";
			for my $y (1..2){
			my $result;
explain									"Running cycle: $y";
			my $x = 0;
			while( $x < 20 and (!$result or $result ne 'EOF') ){
				
#~ ###LogSD	my $expose = 15;
#~ ###LogSD	if( $x == $expose and $y == 1 ){
#~ ###LogSD		$operator->add_name_space_bounds( {
#~ ###LogSD			Test =>{
#~ ###LogSD				_get_next_value_cell =>{
#~ ###LogSD					UNBLOCK =>{
#~ ###LogSD						log_file => 'trace',
#~ ###LogSD					},
#~ ###LogSD				},
#~ ###LogSD			},
#~ ###LogSD		} );
#~ ###LogSD	}

#~ ###LogSD	elsif( $x > ($expose +2) and $y > 0 ){
#~ ###LogSD		exit 1;
#~ ###LogSD	}

lives_ok{	$result = $test_instance->_get_next_value_cell }
										"Collecting data from position: $x";
###LogSD	$phone->talk( level => 'debug', message => [ "result at position -$x- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$x], ] );
is_deeply	$result, $answer_ref->[$x++],"..and see if it has good info";
			}
			}
		
explain									"read through all cells in sequence...";
			for my $y (1..2){
			my $result = undef;
explain									"Running cycle: $y";
			my $x = 20;
			while( $x < 105 and (!$result or $result ne 'EOF') ){
			my	$position = $x - 20;
lives_ok{	$result = $test_instance->_get_next_cell }
										"Collecting data from sheet position: $position";
###LogSD	$phone->talk( level => 'trace', message => [ "result at position -$position- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$x], ] );
is_deeply	$result, $answer_ref->[$x++],"..and see if it has good info";
			}
			}

explain									"read row columns through cells in sequence...";
			for my $y (1..2){
explain									"Running cycle: $y";
			my $y_dim = 1;
			my $x = 105;#############################################################
			my	$result = undef;
			while( $x < 203 and (!$result or $result ne 'EOF') ){
			my	$x_dim = 1;
				$result = undef;
			while( $x < 203 and (!$result or ($result ne 'EOR' and $result ne 'EOF')) ){
lives_ok{	$result = $test_instance->_get_col_row( $x_dim, $y_dim ) }
										"Collecting data for column -$x_dim- and row -$y_dim-";
###LogSD	$phone->talk( level => 'trace', message => [ "result for column -$x_dim- and row -$y_dim- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$x], ] );
is_deeply	$result, $answer_ref->[$x++],"..and see if it has good info";
			$x_dim++;
			}
			$y_dim++;
			}
			}

explain									"read rows through sheet in sequence...";
			for my $y (1..2){
explain									"Running cycle: $y";
			my $result = undef;
			my $y_dim = 1;
			my $x = 203;
			while( $x < 218 and (!$result or $result ne 'EOF') ){
lives_ok{	$result = $test_instance->_get_row_all( $y_dim ) }
										"Collecting data for row -$y_dim-";
###LogSD	$phone->talk( level => 'trace', message => [ "result for row -$y_dim- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$x], ] );
is_deeply	$result, $answer_ref->[$x++],"..and see if it has good info";
			$y_dim++;
			}
			}

lives_ok{
			$workbook_instance->set_empty_is_end( 1 );
			$workbook_instance->set_from_the_edge( 0 );
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet->new(
				file				=> $test_file,
				error_inst			=> $error_instance,
				sheet_name			=> 'Sheet3',
				workbook_instance	=> $workbook_instance,
				is_hidden			=> 0,
			###LogSD	log_space	=> 'Test',
			);
###LogSD	$phone->talk( level => 'trace', message =>[ "Loaded new test instance - without the edges" ] );
}										"Build a Worksheet instance with the edges cut off";

explain									"read through cells without edges in sequence...";
			for my $y (1..2){
			my $result = undef;
explain									"Running cycle: $y";
			my $x = 218;
			while( $x < 257 and (!$result or $result ne 'EOF') ){
			my	$position = $x - 218;
lives_ok{	$result = $test_instance->_get_next_cell }
										"Collecting data from sheet position: $position";
###LogSD	$phone->talk( level => 'trace', message => [ "result at position -$position- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$x], ] );
is_deeply	$result, $answer_ref->[$x++],"..and see if it has good info";
			}
			}

explain									"read row columns through cells without edges in sequence...";
			for my $y (1..2){
explain									"Running cycle: $y";
			my $y_dim = 1;
			my $x = 257;
			my	$result = undef;
			while( $x < 306 and (!$result or $result ne 'EOF') ){
			my	$x_dim = 1;
				$result = undef;
			while( $x < 306 and (!$result or ($result ne 'EOR' and $result ne 'EOF')) ){
lives_ok{	$result = $test_instance->_get_col_row( $x_dim, $y_dim ) }
										"Collecting data for column -$x_dim- and row -$y_dim-";
###LogSD	$phone->talk( level => 'trace', message => [ "result for column -$x_dim- and row -$y_dim- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$x], ] );
is_deeply	$result, $answer_ref->[$x++],"..and see if it has good info";
			$x_dim++;
			}
			$y_dim++;
			}
			}

explain									"read rows through sheet without edges in sequence...";
			for my $y (1..2){
explain									"Running cycle: $y";
			my $result = undef;
			my $y_dim = 1;
			my $x = 306;
			while( $x < 321 and (!$result or $result ne 'EOF') ){
lives_ok{	$result = $test_instance->_get_row_all( $y_dim ) }
										"Collecting data for row -$y_dim-";
###LogSD	$phone->talk( level => 'trace', message => [ "result for row -$y_dim- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$x], ] );
is_deeply	$result, $answer_ref->[$x++],"..and see if it has good info";
			$y_dim++;
			}
			}
is			$test_instance->is_sheet_hidden, 0,
										'Check if the sheet is hidden (Not)';
is_deeply	[ $test_instance->is_column_hidden( 1 .. 6 ) ], [ 0, 0, 1, 1, 0, 0 ],
										'Check that the sheet knows which columns are hidden - by number';
is_deeply	[ $test_instance->is_column_hidden( 'A', 'B', 'C', 'D', 'E', 'F' ) ], [ 0, 0, 1, 1, 0, 0 ],
										'Check that the sheet knows which columns are hidden - by letter';
###LogSD		$operator->add_name_space_bounds( {
###LogSD			Test =>{
###LogSD				is_row_hidden =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'trace',
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
is_deeply	[ $test_instance->is_row_hidden( 0 .. 15 ) ], [ undef, undef, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, 0, 0, undef ],
										'Check that the sheet knows which rows are hidden - by number';
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