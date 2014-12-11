#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::FmtDefault          8#########9
#!evn perl
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; }##### $ENV{ Smart_Comments } = '### ####';
$| = 1;

use	Test::Most tests => 393;
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
use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings;
my	$test_file = ( @ARGV ) ? $ARGV[0] : '../../../../test_files/xl/';
	$test_file .= 'styles.xml';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $capture, $x, @answer,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				target_encoding						excel_region
			);
my  		@class_methods = qw(
				get_target_encoding					set_target_encoding
				get_defined_excel_format			total_defined_excel_formats
				get_excel_region					change_output_encoding
				get_defined_excel_format_list		set_defined_excel_format_list
			);
my			$question_list =[
				[ 'Hello World', "It's a mad mad world" ],
				[undef,'1','1.115111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.115111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.115111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.115111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.115111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.115111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.115111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.115111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','2','-0.1','0.03','0.005','0.00004','0.00005'],
				[undef,'1','2','-0.1','0.03','0.005','0.00004','0.00005'],
				[undef,'1','-200','2000','-2000001','2005','-20005','0.000002','-0.00000000004125'],
				[undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222',
					'6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181',
					'12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667',
					'0.000005','-0.00001','0.9999999','0.019','-0.999'],
				[undef,'0.3333333','-1.6666666','2.1666666','-3.8333333','4.1111111','-5.2222222',
					'6.4444444','-7.5555555','8.7777777','-9.8888888','10.09090909','-11.1818181',
					'12.0833333','-13.4166666','0.12345678','-0.125','0.75','-0.0416666666666667',
					'0.000005','-0.00001','0.9999999','0.019','-0.999'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				undef,								undef,
				undef,								undef,
				undef,								undef,
				undef,								undef,
				[undef,'1','1.11511111111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.11511111111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.11511111111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.11511111111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.11511111111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.11511111111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.11511111111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'1','1.11511111111111','-111111111111115','1.5','-1234.567','59','-60'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'7/4/1776 11:00.234 AM','0.112311','60.99112311','1.500112311','55.0000102311','59.112311','60.345112311'],
				[undef,'1','-200','2000','-2000001','2050','-20050','0.0000002','-0.00000000004125'],
				[ 'Hello World', "It's a mad mad world" ],
			];
my			$answer_list =[
				[ 'General', 'Hello World', "It's a mad mad world" ],
				['0',undef,'1','1','-111111111111115','2','-1235','59','-60'],
				['0.00',undef,'1.00','1.12','-111111111111115.00','1.50','-1234.57','59.00','-60.00'],
				['#,##0',undef,'1','1','-111,111,111,111,115','2','-1,235','59','-60'],
				['#,##0.00',undef,'1.00','1.12','-111,111,111,111,115.00','1.50','-1,234.57','59.00','-60.00'],
				['$#,##0_);($#,##0)',undef,'$1','$1','($111,111,111,111,115)','$2','($1,235)','$59','($60)'],
				['$#,##0_);[Red]($#,##0)',undef,'$1','$1','($111,111,111,111,115)','$2','($1,235)','$59','($60)'],
				['$#,##0.00_);($#,##0.00)',undef,'$1.00','$1.12','($111,111,111,111,115.00)','$1.50','($1,234.57)','$59.00','($60.00)'],
				['$#,##0.00_);[Red]($#,##0.00)',undef,'$1.00','$1.12','($111,111,111,111,115.00)','$1.50','($1,234.57)','$59.00','($60.00)'],
				['0%',undef,'100%','200%','-10%','3%','1%','0%','0%'],
				['0.00%',undef,'100.00%','200.00%','-10.00%','3.00%','0.50%','0.00%','0.01%'],
				['0.00E+00',undef,'1.00E+00','-2.00E+02','2.00E+03','-2.00E+06','2.01E+03','-2.00E+04','2.00E-06','-4.13E-11'],
				['# ?/?',undef,'1/3','-1 2/3','2 1/6','-3 5/6','4 1/9','-5 2/9','6 4/9','-7 5/9','8 7/9',
					'-9 8/9','10 1/9','-11 1/6','12 1/9','-13 3/7','1/8','-1/8','3/4','0','0','0','1','0','-1'],
				['# ??/??',undef,'1/3','-1 2/3','2 1/6','-3 5/6','4 1/9','-5 2/9','6 4/9','-7 5/9','8 7/9',
					'-9 8/9','10 1/11','-11 2/11','12 1/12','-13 5/12','10/81','-1/8','3/4','-1/24','0','0','1','1/53','-1'],
				['yyyy-m-d',undef,'1776-7-4','1904-1-1','1904-3-1','1904-1-2','1904-2-25','1904-2-29','1904-3-1'],
				['d-mmm-yy',undef,'4-Jul-76','1-Jan-04','1-Mar-04','2-Jan-04','25-Feb-04','29-Feb-04','1-Mar-04'],
				['d-mmm',undef,'4-Jul','1-Jan','1-Mar','2-Jan','25-Feb','29-Feb','1-Mar'],
				['mmm-yy',undef,'Jul-76','Jan-04','Mar-04','Jan-04','Feb-04','Feb-04','Mar-04'],
				['h:mm AM/PM',undef,'11:00 AM','2:41 AM','11:47 PM','12:00 PM','12:00 AM','2:41 AM','8:16 AM'],
				['h:mm:ss AM/PM',undef,'11:00:00 AM','2:41:44 AM','11:47:13 PM','12:00:10 PM','12:00:01 AM','2:41:44 AM','8:16:58 AM'],
				['h:mm',undef,'11:00','2:41','23:47','12:00','0:00','2:41','8:16'],
				['h:mm:ss',undef,'11:00:00','2:41:44','23:47:13','12:00:10','0:00:01','2:41:44','8:16:58'],
				['m-d-yy h:mm',undef,'7-4-76 11:00','1-1-04 2:41','3-1-04 23:47','1-2-04 12:00','2-25-04 0:00','2-29-04 2:41','3-1-04 8:16'],
				undef,								undef,
				undef,								undef,
				undef,								undef,
				undef,								undef,
				['#,##0_);(#,##0)',undef,'1','1','(111,111,111,111,115)','2','(1,235)','59','(60)'],
				['#,##0_);[Red](#,##0)',undef,'1','1','(111,111,111,111,115)','2','(1,235)','59','(60)'],
				['#,##0.00_);(#,##0.00)',undef,'1.00','1.12','(111,111,111,111,115.00)','1.50','(1,234.57)','59.00','(60.00)'],
				['#,##0.00_);[Red](#,##0.00)',undef,'1.00','1.12','(111,111,111,111,115.00)','1.50','(1,234.57)','59.00','(60.00)'],
				['_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)','-','1','1','(111,111,111,111,115)','2','(1,235)','59','(60)'],
				['_($*#,##0_);_($*(#,##0);_($*"-"_);_(@_)','$-','$1','$1','$(111,111,111,111,115)','$2','$(1,235)','$59','$(60)'],
				['_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)','-','1.00','1.12','(111,111,111,111,115.00)','1.50','(1,234.57)','59.00','(60.00)'],
				['_($*#,##0.00_);_($*(#,##0.00);_($*"-"??_);_(@_)','$-','$1.00','$1.12','$(111,111,111,111,115.00)','$1.50','$(1,234.57)','$59.00','$(60.00)'],
				['mm:ss',undef,'00:00','41:44','47:13','00:10','00:01','41:44','16:58'],
				['[h]:mm:ss',undef,'-1117548:59:59','2:41:44','1463:47:13','36:00:10','1320:00:01','1418:41:44','1448:16:58'],
				['mm:ss.0',undef,'00:00.2','41:43.7','47:13.0','00:09.7','00:00.9','41:43.7','16:57.7'],
				['##0.0E+0',undef,'1.0E+0','-200.0E+0','2.0E+3','-2.0E+6','2.1E+3','-20.1E+3','200.0E-9','-41.3E-12'],
				[ '@', 'Hello World', "It's a mad mad world" ],
			];
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
lives_ok{
			$test_instance	=	build_instance(
									package	=> 'FmtDefaultTest',
									roles	=>[ 
										'Spreadsheet::XLSX::Reader::LibXML::LogSpace'
									],
									add_roles_in_sequence =>[
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
										"Check that Spreadsheet::XLSX::Reader::LibXML::FmtDefault has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		$test_instance, $_,
} 			@class_methods;
###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
			no warnings 'uninitialized';
			for my $position ( 0 .. ($test_instance->total_defined_excel_formats - 1) ){
			if( $answer_list->[$position] ){
is			$test_instance->get_defined_excel_format( $position ), $answer_list->[$position]->[0],
										,"Check that excel default position -$position- contains: $answer_list->[$position]->[0]";
ok			my $coercion = $test_instance->parse_excel_format_string( $test_instance->get_defined_excel_format( $position ) ),
										,"..and try to turn it into a Type::Tiny coercion";
			for my $row_pos ( 1 .. $#{$answer_list->[$position]} ){
###LogSD	if( $position == 42 and $row_pos == 8 ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			UNBLOCK =>{
###LogSD				log_file => 'warn',
###LogSD			},
###LogSD			Test =>{
###LogSD				_build_number =>{
###LogSD					_build_elements =>{
###LogSD						UNBLOCK =>{
###LogSD							log_file => 'trace',
###LogSD						},
###LogSD						_split_decimal_integer =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD						},
###LogSD						_move_decimal_point =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD						},
###LogSD						_round_decimal =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD						},
###LogSD					},
###LogSD				},
###LogSD				change_output_encoding =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				parse_excel_format_string =>{
###LogSD					UNBLOCK =>{
###LogSD						log_file => 'warn',
###LogSD					},
###LogSD				},
###LogSD				_util_function =>{
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
###LogSD	}elsif( $position == 43 ){
###LogSD		exit 1;
###LogSD	}
###LogSD		$phone->talk( level => 'info', message => [ "Group position: $position", "Test position: $row_pos" ] );
is			$coercion->assert_coerce( $question_list->[$position]->[$row_pos - 1] ), $answer_list->[$position]->[$row_pos],
										,"Testing the excel default coercion -$position- to see if |$question_list->[$position]->[$row_pos - 1]|" . 
											" coerces to: $answer_list->[$position]->[$row_pos]";
			} } }
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