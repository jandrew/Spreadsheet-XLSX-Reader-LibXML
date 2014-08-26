#########1 Test File for Spreadsheet::XLSX::Reader::XMLReader::Worksheet        8#########9
#!env perl
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; };
$| = 1;

use	Test::Most tests => 642;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use DateTime::Format::Flexible;
use Data::Dumper;
use Types::Standard qw(
		InstanceOf
		Str
		Num
	);
use Type::Tiny;
use Type::Coercion;
use	lib
		'../../../../../../Log-Shiras/lib',
		'../../../../../lib',;
use Log::Shiras::Switchboard qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD							main =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'info',
###LogSD								},
###LogSD							},
###LogSD							Test =>{
###LogSD								parse_column_row =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								get_used_position =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								get_excel_position =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'trace',
###LogSD								},
###LogSD								Cell =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'debug',
###LogSD									},
###LogSD								},
###LogSD								Styles =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								get_cell =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'trace',
###LogSD									},
###LogSD								},
###LogSD								_get_column_row =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								_load_unique_bits =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'debug',
###LogSD									},
###LogSD								},
###LogSD								SharedStrings =>{
###LogSD									_set_file_name =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD									_load_unique_bits =>{
###LogSD										UNBLOCK =>{
###LogSD											log_file => 'warn',
###LogSD										},
###LogSD									},
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'debug',
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
use	Spreadsheet::XLSX::Reader::XMLReader::SharedStrings;
use	Spreadsheet::XLSX::Reader::XMLDOM::Styles;

my	$test_file	= ( @ARGV ) ? $ARGV[0] : '../../../../test_files/xl/worksheets/';
my	$shared_strings_file = $test_file . '../sharedStrings.xml';
my	$styles_file = $test_file . '../styles.xml';
	$test_file .= 'sheet3.xml';
my	$log_space	= 'Test';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance, $shared_strings_instance,
			$cell, $styles_instance, $coercion_name, $coercion_to, $coercion_from,
			$custom_coercion
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file_name
				log_space
				sheet_min_col
				sheet_min_row
				sheet_max_col
				sheet_max_row
				sheet_id
				sheet_name
				error_inst
				custom_formats
			);
my  		@instance_methods = qw(
				get_cell
				row_range
				col_range
				error
				clear_error
				set_warnings
				if_warn
				min_col
				has_min_col
				min_row
				has_min_row
				max_col
				has_max_col
				max_row
				has_max_row
				name
				has_custom_format
				get_custom_format
				set_custom_format
				set_custom_formats
			);
my			$answer_ref = [
				[ ['Hello', 'Hello', ],undef,undef,[ 'my', 'my', ],undef,undef,],
				[ undef,undef,undef,undef,undef,undef,],
				[ undef,undef,[ 'World', 'World', ],undef,undef,undef,],
				[ undef,undef,undef,undef,undef,undef,],
				[ [ 'Hello World', 'Hello World', undef, [
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
															], 'A5:B5',],[ '', '', undef, undef,  'A5:B5', ], undef,undef,undef,undef,],
				[ undef,[ 69, 69, ],undef,undef,undef,undef,],
				[ undef,[ 27, 27, ],undef,undef,[ 37145, '12-Sep-05', ],undef,],
				[ undef,[ 42, 42, 'B6-B7' ],undef,undef,undef,undef,],
				[ undef,undef,undef,[ '', '', ],[ '2/6/2011', '2011-02-06T00:00:00', ],[ '2/6/2011', '2011-02-06', ],],
				[ [ 2.1345678901, 2.13, ],undef,undef,undef,undef,undef,],
				[ undef,undef,undef,[ 39118, '6-Feb-11', 'DATEVALUE(E9)', undef, 'D11:E11' ],['', '', undef,undef, 'D11:E11'],undef,],
				[ undef,undef,undef,undef,undef,undef,],
				[ undef,undef,[ ' ', ' ' ],[ 39118, '2011-02-06', 'D11' ],[ 39118, '6-Feb-11', 'D13' ],undef,],
			];

###LogSD	$phone->talk( level => 'info', message => [ "Set up a count from 0 instance" ] );
my	@args_list	= ( system_type => 'apple_excel' );
my	$converter	= DateTimeX::Format::Excel->new( @args_list );
my	$string_via	= sub{ 
						my	$str = $_[0];
						return DateTime::Format::Flexible->parse_datetime( $str );#my	$dt	= 
						#~ return $dt->format_cldr( 'yyyy-M-d' );
					};
my	$num_via	= sub{
						my	$num = $_[0];
						return $converter->parse_datetime( $num );#my	$dt = 
						#~ return $dt->format_cldr( 'yyyy-M-d' );
					};
my	$date_time_from_value = Type::Coercion->new(
	type_coercion_map => [ Num, $num_via, Str, $string_via, ],
);
my	$date_time_type = Type::Tiny->new(
		constrant	=> sub{ ref($_) eq 'DateTime' },
		coercion	=> $date_time_from_value,
	);
###LogSD	$phone->talk( level => 'trace', message => [
###LogSD		"Check coercion:", $date_time_from_value->coerce( '2/6/2011' ) ] );
my	$string_type = Type::Tiny->new(
		name		=> 'YYYYMMDD',
		constraint	=> sub{
			!$_ or (
			$_ =~ /^\d{4}\-(\d{2})-(\d{2})$/ and
			$1 > 0 and $1 < 13 and $2 > 0 and $2 < 32 )
		},
		coercion	=> Type::Coercion->new(
			type_coercion_map =>[
				$date_time_type->coercibles, sub{ my $tmp = $date_time_type->coerce( $_ ); $tmp->format_cldr( 'yyyy-MM-dd' ) },
			],
		),
);
###LogSD	$phone->talk( level => 'trace', message => [
###LogSD		"Check deep coercion:", $string_type->coerce( '2/6/2011' ) ] );

###LogSD	$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::XMLReader::Worksheet', $_,
										"Check that Spreadsheet::XLSX::Reader::XMLReader::Worksheet has the -$_- attribute"
} 			@class_attributes;

#~ lives_ok{
			$error_instance = Spreadsheet::XLSX::Reader::Error->new( should_warn => 0 );
			$styles_instance = Spreadsheet::XLSX::Reader::XMLDOM::Styles->new(
				file_name	=> $styles_file,
				log_space	=> 'Test::Styles',
				error_inst	=> $error_instance,
				epoch_year	=> 1904,
			);
			$shared_strings_instance = Spreadsheet::XLSX::Reader::XMLReader::SharedStrings->new(
				file_name	=> $shared_strings_file,
				log_space	=> 'Test::SharedStrings',
				error_inst	=> $error_instance,
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
				sheet_name				=> 'Sheet3',
			);
			###LogSD	$phone->talk( level => 'info', message =>[ "Loaded test instance" ] );
			$test_instance->set_custom_formats( {
				E9	=> $date_time_type,
				9	=> $string_type,
				D13	=> $string_type,
			} );
#~ }										"Prep a new Worksheet instance";
###LogSD		$phone->talk( level => 'debug', message => [ "Max row is:" . $test_instance->max_row ] );
map{
can_ok		$test_instance, $_,
} 			@instance_methods;
is			$test_instance->get_file_name, $test_file,
										"check that it knows the file name";
is			$test_instance->get_log_space, $log_space,
										"check that it knows the log_space";
is			$test_instance->get_core_element, 'row',
										"check that it knows we are stepping first by row";
is			$test_instance->min_row, 0,
										"check that it knows what the lowest row number is";
is			$test_instance->min_col, 0,
										"check that it knows what the lowest column number is";
is			$test_instance->max_row, 12,
										"check that it knows what the highest row number is";
is			$test_instance->max_col, 5,
										"check that it knows what the highest column number is";
is_deeply	[$test_instance->row_range], [0,12],
										"check for a correct row range";
is_deeply	[$test_instance->col_range], [0,5],
										"check for a correct column range";
										
###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
			#~ for my $row ($test_instance->min_row .. $test_instance->max_row){
			#~ for my $col ($test_instance->min_col .. $test_instance->max_col){
			for my $row (0 .. 12){
			for my $col (0 .. 5){
###LogSD		$phone->talk( level => 'info', message => [ "Checking row: $row", "   and column: $col" ] );
lives_ok{ 
            $cell = $test_instance->get_cell( $row, $col );
}										"Get cell row -$row- and column -$col-";
			if( ref $answer_ref->[$row]->[$col] eq 'ARRAY' ){
###LogSD		$phone->talk( level => 'debug', message => [ "Found data ..." ] );
ok			defined( $cell ),			"Was data found where expected in row -$row- and column -$col-";
is			$cell->unformatted(), $answer_ref->[$row]->[$col]->[0],
										"Check for the expected 'unformatted' value: " . $answer_ref->[$row]->[$col]->[0];
is			$cell->value(), $answer_ref->[$row]->[$col]->[1],
										"Check for the expected 'value': " . $answer_ref->[$row]->[$col]->[1];
			if( $cell->has_formula ){
is			$cell->formula(), $answer_ref->[$row]->[$col]->[2],
										"Check that the formula in the cell is correct: $answer_ref->[$row]->[$col]->[2]";
			}elsif( $answer_ref->[$row]->[$col]->[2] ){
fail									"No formula found where expected for row -$row- and col -$col-";
			}
			if( $cell->has_rich_text ){
				
is_deeply	$cell->get_rich_text(), $answer_ref->[$row]->[$col]->[3],
										"Check that the rich text in the cell is correct: " . Dumper( $cell->get_rich_text() );#$answer_ref->[$row]->[$col]->[3]
			}elsif( $answer_ref->[$row]->[$col]->[3] ){
fail									"No rich_text found where expected for row -$row- and col -$col-";
			}
			if( $cell->is_merged ){
				
is_deeply	$cell->get_merge_range(), $answer_ref->[$row]->[$col]->[4],
										"Check that the merge range in the cell is correct: " . $answer_ref->[$row]->[$col]->[4];
			}elsif( $answer_ref->[$row]->[$col]->[4] ){
fail									"the cell is merged but doesn't know it for row -$row- and col -$col-";
			}
			}else{
is			$cell, $answer_ref->[$row]->[$col],	"Find a null cell where expected at row -$row- and column -$col-";
			}
			}
			}

###LogSD		$phone->talk( level => 'info', message => [ "iterate through using the 'next' functionality ..." ] );
			my	$cell_position = 0;
			for my $row ($test_instance->min_row .. $test_instance->max_row){
			for my $col ($test_instance->min_col .. $test_instance->max_col){
###LogSD		$phone->talk( level => 'info', message => [ "Checking row: $row", "   and column: $col" ] );
lives_ok{
			if( $cell_position == 0 ){
				$cell = $test_instance->get_cell(0,0);
			}else{
				$cell = $test_instance->get_cell;
			}
}										"Get the 'next' cell";
			if( ref $answer_ref->[$row]->[$col] eq 'ARRAY' ){
###LogSD		$phone->talk( level => 'debug', message => [ "Found data ..." ] );
ok			defined( $cell ),			"Was data found where expected at position -$cell_position-";
is			$cell->unformatted(), $answer_ref->[$row]->[$col]->[0],
										"Check for the expected 'unformatted' value: " . $answer_ref->[$row]->[$col]->[0];
is			$cell->value(), $answer_ref->[$row]->[$col]->[1],
										"Check for the expected 'value': " . $answer_ref->[$row]->[$col]->[1];
			if( $cell->has_formula ){
is			$cell->formula(), $answer_ref->[$row]->[$col]->[2],
										"Check that the formula in the cell is correct: $answer_ref->[$row]->[$col]->[2]";
			}elsif( $answer_ref->[$row]->[$col]->[2] ){
fail									"No formula found where expected for position -$cell_position-";
			}
			}else{
is			$cell, $answer_ref->[$row]->[$col],	"Find a null cell where expected at position -$cell_position-";
			}
			$cell_position++;
}
}

###LogSD		$phone->talk( level => 'info', message => [ "Iterate through using 'count from 1' functionality ..." ] );
lives_ok{
			$test_instance = Spreadsheet::XLSX::Reader::XMLReader::Worksheet->new(
				file_name				=> $test_file,
				log_space				=> 'Test',
				epoch_year				=> 1904,
				error_inst				=> $error_instance,
				shared_strings_instance => $shared_strings_instance,
				styles_instance			=> $styles_instance,
				count_from_zero 		=> 0,
			);
			$test_instance->set_custom_formats( {
				E9	=> $date_time_type,
				9	=> $string_type,
				D13	=> $string_type,
			} );
}										"Prep a 'count from one' Worksheet instance";
###LogSD		$phone->talk( level => 'info', message => [ "Max row is:" . $test_instance->max_row ] );
is			$test_instance->min_row, 1,
										"check that it knows what the lowest row number is";
is			$test_instance->min_col, 1,
										"check that it knows what the lowest column number is";
is			$test_instance->max_row, 13,
										"check that it knows what the highest row number is";
is			$test_instance->max_col, 6,
										"check that it knows what the highest column number is";
is_deeply	[$test_instance->row_range], [1,13],
										"check for a correct row range";
is_deeply	[$test_instance->col_range], [1,6],
										"check for a correct column range";
										
###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
			for my $row ($test_instance->min_row .. $test_instance->max_row){
			for my $col ($test_instance->min_col .. $test_instance->max_col){
###LogSD		$phone->talk( level => 'info', message => [ "Checking row: $row", "   and column: $col" ] );
lives_ok{ 
            $cell = $test_instance->get_cell( $row, $col );
}										"Get cell row -$row- and column -$col-";
			if( ref $answer_ref->[$row-1]->[$col-1] eq 'ARRAY' ){
###LogSD		$phone->talk( level => 'debug', message => [ "Found data ..." ] );
ok			defined( $cell ),			"Was data found where expected in row -$row- and column -$col-";
is			$cell->unformatted(), $answer_ref->[$row-1]->[$col-1]->[0],
										"Check for the expected 'unformatted' value: " . $answer_ref->[$row-1]->[$col-1]->[0];
is			$cell->value(), $answer_ref->[$row-1]->[$col-1]->[1],
										"Check for the expected 'value': " . $answer_ref->[$row-1]->[$col-1]->[1];
			if( $cell->has_formula ){
is			$cell->formula(), $answer_ref->[$row-1]->[$col-1]->[2],
										"Check that the formula in the cell is correct: " . $answer_ref->[$row-1]->[$col-1]->[2];
			}elsif( $answer_ref->[$row-1]->[$col-1]->[2] ){
fail									"No formula found where expected for row -$row- and col -$col-";
			}
			}else{
is			$cell, $answer_ref->[$row-1]->[$col-1],	"Find a null cell where expected at row -$row- and column -$col-";
			}
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