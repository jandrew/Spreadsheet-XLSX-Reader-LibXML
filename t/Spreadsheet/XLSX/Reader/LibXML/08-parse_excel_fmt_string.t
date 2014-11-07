#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings    ###9
#!evn perl
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; }
$| = 1;

use	Test::Most tests => 249;
use	Test::Moose;
use Data::Dumper;
use	MooseX::ShortCut::BuildInstance v1.8 qw( build_instance );#
use	lib
		'../../../../../../Log-Shiras/lib',
		'../../../../../lib',;
#~ use Log::Shiras::Switchboard qw( :debug );
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings v0.5;
my	$test_file = ( @ARGV ) ? $ARGV[0] : '../../../../test_files/xl/';
	$test_file .= 'styles.xml';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $capture, $x, @answer,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				epoch_year							cache_formats
				datetime_dates
			);
my  		@class_methods = qw(
				new									get_epoch_year
				get_cache_behavior					get_date_behavior
				set_date_behavior					parse_excel_format_string
				get_defined_excel_format			total_defined_excel_formats
				change_output_encoding				get_excel_region
				has_cached_format					get_cached_format
				set_cached_format
			);
my			$question_list =[
				['[$-409]d-mmm-yy;@',undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				['[$-409]dddd, mmmm dd, yyyy;@',undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				['#,##0E+0',undef,'1','-200','2000','-2000001','2050','-20050','0.0000002','-0.00000000004125'],
				['# ???/???',undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222','6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181','12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667','0.000005','-0.00001','0.9999999','0.019','-0.999'],
				['# ?/2',undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222','6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181','12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667','0.000005','-0.00001','0.9999999','0.019','-0.999'],
				['# ?/4',undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222','6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181','12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667','0.000005','-0.00001','0.9999999','0.019','-0.999'],
				['# ?/8',undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222','6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181','12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667','0.000005','-0.00001','0.9999999','0.019','-0.999'],
				['# ??/16',undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222','6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181','12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667','0.000005','-0.00001','0.9999999','0.019','-0.999'],
				['# ??/10',undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222','6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181','12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667','0.000005','-0.00001','0.9999999','0.019','-0.999'],
				['# ???/100',undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222','6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181','12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667','0.000005','-0.00001','0.9999999','0.019','-0.999'],
				['# ??????/??????',undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222','6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181','12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667','0.000005','-0.00001','0.9999999','0.019','-0.999'],
				];
my			$answer_list =[
				['[$-409]d-mmm-yy;@',undef,'4-Jul-76','1-Jan-04','1-Mar-04','2-Jan-04','25-Feb-04','29-Feb-04','1-Mar-04'],
				['[$-409]dddd, mmmm dd, yyyy;@',undef,'Thursday, July 04, 1776','Friday, January 01, 1904','Tuesday, March 01, 1904','Saturday, January 02, 1904','Thursday, February 25, 1904','Monday, February 29, 1904','Tuesday, March 01, 1904'],
				['#,##0E+0',undef,'1E+0','-200E+0','2,000E+0','-200E+4','2,050E+0','-2E+4','20E-8','-41E-12'],
				['# ???/???',undef,'1/3','-1 2/3','2 1/6','-3 5/6','4 1/9','-5 2/9','6 4/9','-7 5/9','8 7/9','-9 8/9','10 1/11','-11 2/11','12 1/12','-13 5/12','10/81','-1/8','3/4','-1/24','0','0','1','11/579','-998/999'],
				['# ?/2',undef,'1/2','-1 1/2','2','-4','4','-5','6 1/2','-7 1/2','9','-10','10','-11','12','-13 1/2','0','0','1','0','0','0','1','0','-1'],
				['# ?/4',undef,'1/4','-1 3/4','2 1/4','-3 3/4','4','-5 1/4','6 2/4','-7 2/4','8 3/4','-10','10','-11 1/4','12','-13 2/4','0','-1/4','3/4','0','0','0','1','0','-1'],
				['# ?/8',undef,'3/8','-1 5/8','2 1/8','-3 7/8','4 1/8','-5 2/8','6 4/8','-7 4/8','8 6/8','-9 7/8','10 1/8','-11 1/8','12 1/8','-13 3/8','1/8','-1/8','6/8','0','0','0','1','0','-1'],
				['# ??/16',undef,'5/16','-1 11/16','2 3/16','-3 13/16','4 2/16','-5 4/16','6 7/16','-7 9/16','8 12/16','-9 14/16','10 1/16','-11 3/16','12 1/16','-13 7/16','2/16','-2/16','12/16','-1/16','0','0','1','0','-1'],
				['# ??/10',undef,'3/10','-1 7/10','2 2/10','-3 8/10','4 1/10','-5 2/10','6 4/10','-7 6/10','8 8/10','-9 9/10','10 1/10','-11 2/10','12 1/10','-13 4/10','1/10','-1/10','8/10','0','0','0','1','0','-1'],
				['# ???/100',undef,'33/100','-1 67/100','2 17/100','-3 83/100','4 11/100','-5 22/100','6 44/100','-7 56/100','8 78/100','-9 89/100','10 9/100','-11 18/100','12 8/100','-13 42/100','12/100','-13/100','75/100','-4/100','0','0','1','2/100','-1'],
				['# ??????/??????',undef,'1/3','-1 2/3','2 1/6','-3 5/6','4 1/9','-5 2/9','6 4/9','-7 5/9','8 7/9','-9 8/9','10 1/11','-11 2/11','12 1/12','-13 5/12','10/81','-1/8','3/4','-1/24','1/200000','-1/100000','1','19/1000','-999/1000'],
			];
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
lives_ok{
			$test_instance	=	build_instance(
									package	=> 'ParseExcelFormatStringsTest4',
									roles	=>[ 
										'Spreadsheet::XLSX::Reader::LibXML::LogSpace'
									],
									add_roles_in_sequence =>[
										'Spreadsheet::XLSX::Reader::LibXML::UtilFunctions',
										'Spreadsheet::XLSX::Reader::LibXML::FmtDefault',
										'Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings'
									],
									log_space				=> 'Test',
									epoch_year				=> 1904,
								);
}										"Prep a test ParseExcelFormatStrings instance";
map{ 
has_attribute_ok
			$test_instance, $_,
										"Check that Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		$test_instance, $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'trace', message => [ 'Test instance:', $test_instance ] );
###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
			no warnings 'uninitialized';
			for my $position ( 0 .. $#$question_list ){
###LogSD	if( $position == 2 ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			main =>{
###LogSD				UNBLOCK =>{
###LogSD					log_file => 'debug',
###LogSD				},
###LogSD			},
###LogSD			Test =>{
###LogSD				_build_number =>{
###LogSD					_build_scientific_sub =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD				change_output_encoding =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				_util_function =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'trace',
###LogSD					},
###LogSD					_gcd =>{
###LogSD						BLOCK=>{
###LogSD							log_file => 'fatal',
###LogSD						},
###LogSD					},
###LogSD					_integer_and_decimal =>{
###LogSD						BLOCK=>{
###LogSD							log_file => 'fatal',
###LogSD						},
###LogSD					},
###LogSD					_best_fraction =>{
###LogSD						BLOCK=>{
###LogSD							log_file => 'fatal',
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD			},
###LogSD		} );
###LogSD	}elsif( $position == 3 ){
###LogSD		exit 1;
###LogSD	}
###LogSD		$phone->talk( level => 'debug', message => [ 'processing excel format string: ' . $question_list->[$position]->[0]  ] );
ok			my $coercion = $test_instance->parse_excel_format_string( $question_list->[$position]->[0] ),
										"Build a coercion with excel format string: $question_list->[$position]->[0]";
###LogSD		$phone->talk( level => 'debug', message => [ 'Built a coercion named : ' . $coercion->name  ] );
			for my $row_pos ( 1 .. $#{$question_list->[$position]} ){
###LogSD		$phone->talk( level => 'info', message => [ "Group position: $position", "Test position: $row_pos" ] );
###LogSD		$phone->talk( level => 'debug', message => [ "Attempting to coerce: $question_list->[$position]->[$row_pos]"  ] );
is			$coercion->assert_coerce( $question_list->[$position]->[$row_pos] ), $answer_list->[$position]->[$row_pos],
										"Testing the coercion for -$question_list->[$position]->[0]- to see if " .
											"|$question_list->[$position]->[$row_pos]|" . 
											" coerces to: $answer_list->[$position]->[$row_pos]";
			} }
ok			$test_instance->set_date_behavior( 1 ),
										"Set the date output to privide DateTime objects rather than strings";
			my $date_string = 'yyyy-mm-dd';
			my $time		= 55.0000102311;
ok			my $coercion = $test_instance->parse_excel_format_string( $date_string ),
										"Build a coercion with excel format string: $date_string";
is			ref $coercion->assert_coerce( $time ), 'DateTime',
										"Checking that a DateTime object was returned";
is			$coercion->assert_coerce( $time ), '1904-02-25T00:00:01',
										"Checking that the date and time are correct: 1904-02-25T00:00:01";
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