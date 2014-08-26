#########1 Test File for Spreadsheet::XLSX::Reader::XMLDOM::Cell      7#########8#########9
#!env perl
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; }
$| = 1;

use	Test::Most tests => 73;
use	Test::Moose;
use	Data::Dumper;
use	MooseX::ShortCut::BuildInstance v1.8 qw( build_instance );
use	lib
		'../../../../../Log-Shiras/lib',
		'../../../../lib',;
#~ use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::Cell;
use	Spreadsheet::XLSX::Reader::Error;
use	Spreadsheet::XLSX::Reader::Types qw( PassThroughType ZeroFromNum FourteenFromWinExcelNum );
my	$test_file = ( @ARGV ) ? $ARGV[0] : '../../../../test_files/';
	$test_file .= 'styles.xml';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $capture, $x, @answer,# $error_instance,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				value_encoding
				value_type
				cell_column
				cell_row
				raw_value
				cell_formula
				merge_range
				rich_text
				font
				fill
				borderId
				fillId
				fontId
				applyFont
				applyNumberFormat
				border
				alignment
				numFmtId
				xfId
				NumberFormat
				error_inst
				log_space
			);
my  		@class_methods = qw(
				new
				get_log_space
				set_log_space
				unformatted
				value
				is_not_empty
				
				encoding
				type
				column
				row
				formula
				has_formula
				get_merge_range
				is_merged
				get_rich_text
				has_rich_text
				get_font
				get_fill
				get_border
				get_alignment
				get_format
				has_format
				clear_format
				set_format
				format_name
				error
				set_error
				clear_error
				set_warnings
				if_warn
				cell_id
			);
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::Cell', $_,
										"Check that Spreadsheet::XLSX::Reader::Cell has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		'Spreadsheet::XLSX::Reader::Cell', $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$test_instance	=	Spreadsheet::XLSX::Reader::Cell->new(
									'NumberFormat' => PassThroughType->plus_coercions( ZeroFromNum ),
									'font' => {
										'color' => {
											'theme' => '1'
										},
										'sz' => '11',
										'name' => 'Calibri',
										'scheme' => 'minor',
										'family' => '2'
									},
									'numFmtId' => 0,
									'fillId' => 0,
									'fill' => {
										'patternFill' => {
											'patternType' => 'none'
										}
									},
									'xfId' => '0',
									'alignment' => {
										'horizontal' => 'left'
									},
									'merge_range' => 'A5:B5',
									'cell_formula' => 'A1&" "&C3',
									'rich_text' => [
										[
											2,
											{
												'color' => {
													'rgb' => 'FFFF0000'
												},
												'sz' => '11',
												'b' => 1,
												'scheme' => 'minor',
												'rFont' => 'Calibri',
												'family' => '2'
											}
										],
										[
											6,
											{
												'color' => {
													'rgb' => 'FF0070C0'
												},
												'sz' => '20',
												'b' => 1,
												'scheme' => 'minor',
												'rFont' => 'Calibri',
												'family' => '2'
											}
										]
									],
									'fontId' => 0,
									'error_inst' => bless( {
														'should_warn' => 0,
															'log_space' => 'Spreadsheet::XLSX::Reader::LogSpace'
													}, 'Spreadsheet::XLSX::Reader::Error' ),
									'value_encoding' => 'UTF-8',
									'raw_value' => 'Hello World',
									'log_space' => 'Test::Cell',
									'cell_column' => 0,
									'cell_row' => 4,
									'border' => {
										'left' => undef,
										'right' => undef,
										'diagonal' => undef,
										'top' => undef,
										'bottom' => undef
									},
									'value_type' => 's',
									'borderId' => 0,
								);
}										"Prep a new Cell instance";

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
is			$test_instance->formula, 'A1&" "&C3',
										"Check that the 'formula' method returns: A1&\" \"&C3";
ok			$test_instance->is_merged,	"Check that the 'is_merged' method returns TRUE";
is_deeply	$test_instance->get_merge_range, 'A5:B5',
										"Check that the get_merge_range method returns: A5:B5";
is_deeply	$test_instance->get_merge_range( 'array' ), [[0,4], [1,4]],
										"Check that the 'get_merge_range( 'array' )' method returns: [[0,4], [1,4]]";
is			$test_instance->encoding, 'UTF-8',
										"Check that the 'encoding' method returns: UTF-8";
is			$test_instance->row, 4,
										"Check that the 'row' method returns: 4";
is			$test_instance->column, 0,
										"Check that the 'column' method returns: 0";
is			$test_instance->unformatted, 'Hello World',
										"Check that the 'unformatted' method returns: Hello World";
is			$test_instance->type, 's',
										"Check that the 'type' method returns: s";
is			$test_instance->cell_id, 'A5',
										"Check that the 'cell_id' method returns: A5";
is			$test_instance->has_format, 1,
										"Check that the 'has_number_format' method returns: TRUE";
is			$test_instance->format_name, 'PassThroughType',
										"Check that the 'format_name' method returns: PassThroughType";
is			$test_instance->get_format->display_name, 'PassThroughType',
										"Get the full Type::Coercion object and call a Type::Coercion method on it returning: PassThroughType";
is			$test_instance->value, 'Hello World',
										"Pull the formatted value and see that it is correct";
lives_ok{	$test_instance->clear_format }
										"Clear the format";
is			$test_instance->has_format, '',
										"... and check that the 'has_format' method returns: FALSE";
is			$test_instance->set_format( FourteenFromWinExcelNum ), FourteenFromWinExcelNum,
										"Set the format object to: FourteenFromWinExcelNum";
is			$test_instance->format_name, 'FourteenFromWinExcelNum',
										"... and check that the 'format_name' method returns: FourteenFromWinExcelNum";
is			$test_instance->unformatted, 'Hello World',
										"Pull the unformatted value and check that it is: Hello World";
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
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "name_space - %-50s | level - %-6s |\nfile_name  - %-50s | line  - %04d   |\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{name_space}, $_[0]->{level},
###LogSD					$_[0]->{filename}, $_[0]->{line},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;