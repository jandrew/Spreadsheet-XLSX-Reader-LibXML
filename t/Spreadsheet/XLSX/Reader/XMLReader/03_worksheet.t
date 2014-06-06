#########1 Test File for Spreadsheet::XLSX::Reader::XMLReader::Worksheet        8#########9
#!perl
$| = 1;

use	Test::Most;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use Type::Coercion;
#~ use Type::Utils -all;
#~ use Type::Library
	#~ -base,
	#~ -declare => qw(
		#~ CustomCoercion
	#~ );
use Data::Dumper;
use Types::Standard qw(
		InstanceOf
	);
use	lib
		'../../../../../../Log-Shiras/lib',
		'../../../../../lib',;
use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD							Test =>{
###LogSD								_get_column_row =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								_load_unique_bits =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								SharedStrings =>{
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
use	Spreadsheet::XLSX::Reader::XMLReader::Worksheet;
use	Spreadsheet::XLSX::Reader::Error;
use	Spreadsheet::XLSX::Reader::Types qw( CustomFormat );
use	Spreadsheet::XLSX::Reader::XMLReader::SharedStrings;
use	Spreadsheet::XLSX::Reader::XMLDOM::Styles;

#~ declare_coercion CustomCoercion,
	#~ to_type Num, from Str,
	#~ via{ 
		#~ my $num = $_;
		#~ return '0' if( !defined $num );
		#~ $num =~ /(-?)(\d*)\.?(\d)/;
		#~ my ( $sign, $integer, $first_dec ) = ( $1, $2, $3 );
		#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new(
		#~ ###LogSD				name_space 	=> $log_space . '::OneFromNum', );
		#~ ###LogSD	no warnings 'uninitialized';
		#~ ###LogSD		$phone->talk( level => 'info', message => [
		#~ ###LogSD				"Coercing num: $num",
		#~ ###LogSD				"Into integer using: |$sign| |$integer| |$first_dec|" ] );
		#~ ###LogSD	use warnings 'uninitialized';
		#~ $integer += 1 if( $first_dec and $first_dec > 4);
		#~ my $return = "$sign$integer";
		#~ ###LogSD		$phone->talk( level => 'info', message => [ "Returning: $return" ] );
		#~ return $return;
	#~ };
my	$test_file	= ( @ARGV ) ? $ARGV[0] : '../../../../test_files/xl/worksheets/';
my	$shared_strings_file = $test_file . '../sharedStrings.xml';
my	$styles_file = $test_file . '../styles.xml';
	$test_file .= 'sheet3.xml';
my	$log_space	= 'Test';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance, $shared_strings_instance, $cell, $styles_instance,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file_name
				log_space
				sheet_min_col
				sheet_min_row
				sheet_max_col
				sheet_max_row
				sheet_rel_id
				sheet_id
				sheet_position
				sheet_name
				error_inst
				custom_formats
			);
my  		@instance_methods = qw(
				get_position
				get_file_name
				where_am_i
				has_position
				get_log_space
				set_log_space
				get_core_element
				min_row
				has_min_row
				max_row
				has_max_row
				min_col
				has_min_col
				max_col
				has_max_col
				rel_id
				sheet_id
				position
				name
				row_range
				col_range
				get_cell
				has_custom_format
				get_custom_format
			);
my			$answer_ref = [
				[
					'<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="D11" i="1"/>',
				],
			];
###LogSD	$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::XMLReader::Worksheet', $_,
										"Check that Spreadsheet::XLSX::Reader::XMLReader::Worksheet has the -$_- attribute"
} 			@class_attributes;

###LogSD	$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			my	$custom_coercion = {
					s => 	Type::Coercion->new(
								type_constraint	=> InstanceOf[ 'DateTime' ],
								name	=> 'DateFromString',
							),
				};
			$error_instance = Spreadsheet::XLSX::Reader::Error->new;
			$shared_strings_instance = Spreadsheet::XLSX::Reader::XMLReader::SharedStrings->new(
				file_name	=> $shared_strings_file,
				log_space	=> 'Test::SharedStrings',
				error_inst	=> $error_instance,
			);
			$styles_instance = Spreadsheet::XLSX::Reader::XMLDOM::Styles->new(
				file_name	=> $styles_file,
				log_space	=> 'Test::Styles',
				error_inst	=> $error_instance,
				epoch_year	=> 1904,
			);
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Error instance:", $error_instance, "SharedStrings instance:", $shared_strings_instance ] );
			$test_instance = Spreadsheet::XLSX::Reader::XMLReader::Worksheet->new(
				file_name				=> $test_file,
				log_space				=> 'Test',
				epoch_year				=> 1904,
				error_inst				=> $error_instance,
				shared_strings_instance => $shared_strings_instance,
				styles_instance			=> $styles_instance,
				custom_formats			=> {
					E9 => $custom_coercion,
				},
			);
}										"Prep a new Worksheet instance";
###LogSD		$phone->talk( level => 'info', message => [ "Max row is:" . $test_instance->max_row ] );
map{
can_ok		$test_instance, $_,
} 			@instance_methods;
is			$test_instance->get_file_name, $test_file,
										"check that it knows the file name";
is			$test_instance->get_log_space, $log_space,
										"check that it knows the log_space";
is			$test_instance->get_core_element, 'row',
										"check that it knows we are stepping first by row";
is			$test_instance->min_row, 1,
										"check that it knows what the lowest row number is";
is			$test_instance->min_col, 1,
										"check that it knows what the lowest column number is";
is			$test_instance->max_row, 13,
										"check that it knows what the highest row number is";
is			$test_instance->max_col, 5,
										"check that it knows what the highest column number is";
is_deeply	[$test_instance->row_range], [1,13],
										"check for a correct row range";
is_deeply	[$test_instance->col_range], [1,5],
										"check for a correct column range";
										
###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
			$x = 0;
map{
			my $row = $_;
map{
###LogSD		$phone->talk( level => 'info', message => [ "Checking row: $row", "   and cell: $_" ] );
#~ lives_ok{ 
            $cell = $test_instance->get_cell( $row, $_ );
#~ }										"Make sure you can get a cell";
###LogSD		$phone->talk( level => 'debug', message => [ "Found data ..." ] );
            next unless $cell;
is			$cell->value(), $answer_ref->[$x]->[0],
										"Check for the expected 'value'";
is			$cell->unformatted(), $answer_ref->[$x++]->[1],
										"Check for the expected 'unformatted' value";
}			($test_instance->min_col .. $test_instance->max_col);
}			($test_instance->min_row .. $test_instance->max_row);
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