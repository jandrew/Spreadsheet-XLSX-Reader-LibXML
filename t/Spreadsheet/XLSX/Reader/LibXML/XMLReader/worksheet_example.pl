#!/usr/bin/env perl
$|=1;
use Data::Dumper;
use MooseX::ShortCut::BuildInstance qw( build_instance );
use Types::Standard qw( Bool HasMethods );
use Spreadsheet::XLSX::Reader::LibXML::Error;
use Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet;
my  $error_instance = Spreadsheet::XLSX::Reader::LibXML::Error->new( should_warn => 0 );
my  $workbook_instance = build_instance(
		package     => 'WorkbookInstance',
		add_methods =>{
			counting_from_zero         => sub{ return 0 },
			boundary_flag_setting      => sub{},
			change_boundary_flag       => sub{},
			_has_shared_strings_file   => sub{ return 1 },
			get_shared_string_position => sub{},
			_has_styles_file           => sub{},
			get_format_position        => sub{},
			get_group_return_type      => sub{},
			set_group_return_type      => sub{},
			get_epoch_year             => sub{ return '1904' },
			change_output_encoding     => sub{ $_[0] },
			get_date_behavior          => sub{},
			set_date_behavior          => sub{},
			get_empty_return_type      => sub{ return 'undef_string' },
			get_values_only            => sub{},
			set_values_only            => sub{},
		},
		add_attributes =>{
			error_inst =>{
				isa => 	HasMethods[qw(
							error set_error clear_error set_warnings if_warn
						) ],
				clearer  => '_clear_error_inst',
				reader   => 'get_error_inst',
				required => 1,
				handles =>[ qw(
					error set_error clear_error set_warnings if_warn
				) ],
			},
			empty_is_end =>{
				isa     => Bool,
				writer  => 'set_empty_is_end',
				reader  => 'is_empty_the_end',
				default => 0,
			},
			from_the_edge =>{
				isa     => Bool,
				reader  => '_starts_at_the_edge',
				writer  => 'set_from_the_edge',
				default => 1,
			},
		},
		error_inst => $error_instance,
	);
my $test_instance = Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet->new(
		file              => '../../../../../test_files/xl/worksheets/sheet3.xml',
		error_inst        => $error_instance,
		sheet_name        => 'Sheet3',
		workbook_instance => $workbook_instance,
	);
my $x = 0;
my $result;
while( $x < 20 and (!$result or $result ne 'EOF') ){
	$result = $test_instance->_get_next_value_cell;
	print "Collecting data from position: $x\n" . Dumper( $result ) . "\n";
	$x++;
}