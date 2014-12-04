#########1 Test File for Spreadsheet::XLSX::Reader::XMLReader::Worksheet        8#########9
#!/usr/bin/env perl
my ( $lib, $test_file );
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

use	Test::Most tests => 1326;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use Types::Standard qw( Bool );
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
###LogSD									log_file => 'debug',
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
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet;
use	DateTimeX::Format::Excel;
use	DateTime::Format::Flexible;
use	Type::Coercion;
use	Type::Tiny;

	$test_file	= ( @ARGV ) ? $ARGV[0] : $test_file;
	$test_file .= 'sheet3.xml';
my	$log_space	= 'Test';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $error_instance, $workbook_instance,
	);
my 			@class_attributes = qw(
				file_name
				error_inst
				log_space
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
				get_file_name
				error
				clear_error
				set_warnings
				if_warn
				get_log_space
				parse_column_row
				build_cell_label
				parse_element
				_get_next_value_cell
				_get_next_cell
				_get_col_row
				_get_row_all
			);
my			$answer_ref = [
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's' },
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's' },
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, },
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, },
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2 },
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, },
				{ r => 'D10', row => 10, col => 4, s => 1, },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', },
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, },
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12' },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, },
				'EOF',
				undef, undef, undef, undef, undef, undef,
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				undef, undef,
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's' },
				undef, undef,
				undef, undef, undef, undef, undef, undef,
				undef, undef,
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's' },
				undef, undef, undef,
				undef, undef, undef, undef, undef, undef,
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				undef, undef, undef, undef,
				undef,
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, },
				undef, undef, undef, undef,
				undef,
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, },
				undef, undef,
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2 },
				undef,
				undef,
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, },
				undef, undef, undef, undef,
				undef, undef, undef,
				{ r => 'D10', row => 10, col => 4, s => 1, },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', },
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, },
				undef, undef, undef, undef, undef,
				undef, undef, undef,
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12' },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				undef,
				undef, undef, undef, undef, undef, undef,
				undef, undef,
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, },
				undef,
				'EOF',
				undef, undef, undef, undef, undef, undef,'EOR',
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				undef, undef,
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's' },
				undef, undef,'EOR',
				undef, undef, undef, undef, undef, undef,'EOR',
				undef, undef,
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's' },
				undef, undef, undef,'EOR',
				undef, undef, undef, undef, undef, undef,'EOR',
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				undef, undef, undef, undef,'EOR',
				undef,
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, },
				undef, undef, undef, undef,'EOR',
				undef,
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, },
				undef, undef,
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2 },
				undef,'EOR',
				undef,
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, },
				undef, undef, undef, undef,'EOR',
				undef, undef, undef,
				{ r => 'D10', row => 10, col => 4, s => 1, },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', },
				'EOR',
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, },
				undef, undef, undef, undef, undef,'EOR',
				undef, undef, undef,
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12' },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				undef,'EOR',
				undef, undef, undef, undef, undef, undef,'EOR',
				undef, undef,
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, },
				undef,'EOF',
				[undef, undef, undef, undef, undef, undef,],
				[
					{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },undef, undef,
					{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's' },undef, undef,
				],
				[undef, undef, undef, undef, undef, undef,],
				[
					undef, undef,{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's' }, undef, undef, undef,
				],
				[undef, undef, undef, undef, undef, undef,],
				[
					{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
					{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', }, undef, undef, undef, undef,
				],
				[
					undef,{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, }, undef, undef, undef, undef,
				],
				[
					undef,{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, },undef, undef,
					{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2 }, undef,
				],
				[
					undef,{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, }, undef, undef, undef, undef,
				],
				[
					undef, undef, undef,{ r => 'D10', row => 10, col => 4, s => 1, },
					{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, },
					{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', },
				],
				[
					{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, }, undef, undef, undef, undef, undef,
				],
				[
					undef, undef, undef,{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12' },
					{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', }, undef,
				],
				[undef, undef, undef, undef, undef, undef,],
				[
					undef, undef,{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', },
					{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, },
					{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, }, undef,
				],	
				'EOF',
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				undef, undef,
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's' },
				undef,
				undef, undef,
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's' },
				undef,
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				undef,
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, },
				undef,
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, },
				undef, undef,
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2 },
				undef,
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, },
				undef, undef, undef,
				{ r => 'D10', row => 10, col => 4, s => 1, },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', },
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, },
				undef, undef, undef,
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12' },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				undef,
				undef, undef,
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, },
				'EOF',
				'EOR',
				{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },
				undef, undef,
				{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's' },
				'EOR',
				'EOR',
				undef, undef,
				{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's' },
				'EOR',
				'EOR',
				{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
				{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				'EOR',
				undef,
				{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, },
				'EOR',
				undef,
				{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, },
				undef, undef,
				{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2 },
				'EOR',
				undef,
				{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, },
				'EOR',
				undef, undef, undef,
				{ r => 'D10', row => 10, col => 4, s => 1, },
				{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, },
				{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', },
				'EOR',
				{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, },
				'EOR',
				undef, undef, undef,
				{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12' },
				{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				'EOR',
				'EOR',
				undef, undef,
				{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', },
				{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, },
				{ r => 'E14', row => 14, col => 5, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D14' }, s => 2, },
				'EOF',
				[],
				[
					{ r => 'A2', row => 2, col => 1, v =>{ raw_text => '0' }, t => 's' },undef, undef,
					{ r => 'D2', row => 2, col => 4, v =>{ raw_text => '2' }, t => 's' },
				],
				[],
				[
					undef, undef,{ r => 'C4', row => 4, col => 3, v =>{ raw_text => '1' }, s => '7', t => 's' },
				],
				[],
				[
					{ r => 'A6', row => 6, col => 1, v =>{ raw_text => '15' }, s => '11', t => 's', cell_merge => 'A6:B6' },
					{ r => 'B6', row => 6, col => 2, s => '11', cell_merge => 'A6:B6', },
				],
				[
					undef,{ r => 'B7', row => 7, col => 2, v =>{ raw_text => '69' }, },
				],
				[
					undef,{ r => 'B8', row => 8, col => 2, v =>{ raw_text => '27' }, },undef, undef,
					{ r => 'E8', row => 8, col => 5, v =>{ raw_text => '37145' }, s => 2 },
				],
				[
					undef,{ r => 'B9', row => 9, col => 2, v =>{ raw_text => '42' }, f =>{ raw_text => 'B7-B8' }, },
				],
				[
					undef, undef, undef,{ r => 'D10', row => 10, col => 4, s => 1, },
					{ r => 'E10', row => 10, col => 5, v =>{ raw_text => '14' }, t => 's', s => 6, },
					{ r => 'F10', row => 10, col => 6, v =>{ raw_text => '14' }, s => 2, t => 's', },
				],
				[
					{ r => 'A11', row => 11, col => 1, v =>{ raw_text => '2.1345678901' }, s => 8, },
				],
				[
					undef, undef, undef,{ r => 'D12', row => 12, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'DATEVALUE(E10)' }, s => 10, cell_merge => 'D12:E12' },
					{ r => 'E12', row => 12, col => 5, s => 10, cell_merge => 'D12:E12', },
				],
				[],
				[
					undef, undef,{ r => 'C14', row => 14, col => 3, v =>{ raw_text => '3' }, t => 's', },
					{ r => 'D14', row => 14, col => 4, v =>{ raw_text => '39118' }, f =>{ raw_text => 'D12' }, s => 9, },
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
			$workbook_instance = build_instance(
									package	=> 'WorkbookInstance',
									add_methods =>{
										counting_from_zero			=> sub{ return 0 },
										boundary_flag_setting		=> sub{},
										change_boundary_flag		=> sub{},
										_has_shared_strings_file	=> sub{ return 1 },
										get_shared_string_position	=> sub{},
										_has_styles_file			=> sub{},
										get_format_position			=> sub{},
										get_group_return_type		=> sub{},
										set_group_return_type		=> sub{},
										get_epoch_year				=> sub{ return '1904' },
										change_output_encoding		=> sub{ $_[0] },
										get_date_behavior			=> sub{},
										set_date_behavior			=> sub{},
									},
									add_attributes =>{
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
									},
								);
			$error_instance = Spreadsheet::XLSX::Reader::LibXML::Error->new( should_warn => 0 );
			$test_instance	= Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet->new(
				file_name				=> $test_file,
				log_space				=> 'Test',
				error_inst				=> $error_instance,
				sheet_name				=> 'Sheet3',
				workbook_instance		=> $workbook_instance,
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
}										"Prep a new Worksheet instance";
###LogSD		$phone->talk( level => 'debug', message => [ "Max row is:" . $test_instance->max_row ] );
map{
can_ok		$test_instance, $_,
} 			@instance_methods;
is			$test_instance->get_file_name, $test_file,
										"check that it knows the file name";
is			$test_instance->get_log_space, $log_space,
										"check that it knows the log_space";
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
			while( !$result or $result ne 'EOF' ){
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
			my $x = 19;
			while( $x < 105 and (!$result or $result ne 'EOF') ){
			my	$position = $x - 19;
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
			my $x = 104;
			my	$result = undef;
			while( $x < 202 and (!$result or $result ne 'EOF') ){
			my	$x_dim = 1;
				$result = undef;
			while( $x < 202 and (!$result or ($result ne 'EOR' and $result ne 'EOF')) ){
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
			my $x = 202;
			while( $x < 217 and (!$result or $result ne 'EOF') ){
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
				file_name				=> $test_file,
				log_space				=> 'Test',
				error_inst				=> $error_instance,
				sheet_name				=> 'Sheet3',
				workbook_instance		=> $workbook_instance,
			);
###LogSD	$phone->talk( level => 'trace', message =>[ "Loaded new test instance - without the edges" ] );
}										"Build a Worksheet instance with the edges cut off";

explain									"read through cells without edges in sequence...";
			for my $y (1..2){
			my $result = undef;
explain									"Running cycle: $y";
			my $x = 217;
			while( $x < 258 and (!$result or $result ne 'EOF') ){
			my	$position = $x - 217;
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
			my $x = 256;
			my	$result = undef;
			while( $x < 305 and (!$result or $result ne 'EOF') ){
			my	$x_dim = 1;
				$result = undef;
			while( $x < 305 and (!$result or ($result ne 'EOR' and $result ne 'EOF')) ){
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
			my $x = 305;
			while( $x < 330 and (!$result or $result ne 'EOF') ){
lives_ok{	$result = $test_instance->_get_row_all( $y_dim ) }
										"Collecting data for row -$y_dim-";
###LogSD	$phone->talk( level => 'trace', message => [ "result for row -$y_dim- is:", $result,
###LogSD		'Against answer:', $answer_ref->[$x], ] );
is_deeply	$result, $answer_ref->[$x++],"..and see if it has good info";
			$y_dim++;
			}
			}
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