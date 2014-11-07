#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::Types     7#########8#########9
#!env perl
my ( $lib, $test_file );
BEGIN{
	$ENV{PERL_TYPE_TINY_XS} = 0;
	my	$start_deeper = 1;
	$lib		= 'lib';
	$test_file	= 't/test_files/';
	for my $next ( <*> ){
		if( ($next eq 't') and -d $next ){
			$start_deeper = 0;
			last;
		}
	}
	if( $start_deeper ){
		$lib		= '../../../../../' . $lib;
		$test_file	= '../../../../test_files/';
	}
}
$| = 1;

use	Test::Most tests => 149;
use	Test::Moose;
use Data::Dumper;
use Capture::Tiny qw( capture_stderr );
use	lib 
		'../../../../../../Log-Shiras/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard v0.21 qw( :debug );#
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							Test =>{
###LogSD								UNBLOCK =>{
###LogSD									log_file => 'info',
###LogSD								},
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML::Types v0.1 qw(
		PassThroughType
		FileName
		XLSXFile
		XMLFile
		ParserType
		EpochYear
		
		Excel_number_0
		OneFromNum
		TwoFromNum
		ThreeFromNum
		FourFromNum
		NineFromNum
		TenFromNum
		ElevenFromNum
		TwelveFromNum
		FourteenFromWinExcelNum
		FourteenFromAppleExcelNum
		FifteenFromWinExcelNum
		FifteenFromAppleExcelNum
		SixteenFromWinExcelNum
		SixteenFromAppleExcelNum
		SeventeenFromWinExcelNum
		SeventeenFromAppleExcelNum
		EighteenFromNum
);
no	warnings 'once';
$Spreadsheet::XLSX::Reader::Types::log_space = 'Test';
use	warnings 'once';
my	$test_dir	= ( @ARGV ) ? $ARGV[0] : $test_file;
my	$xlsx_file	= 'TestBook.xlsx';
my	$xml_file	= '[Content_Types].xml';
my  ( 
			$position, $counter, $capture,
	);
my 			$row = 0;
my			$question_ref =[
				$xlsx_file,
				$xml_file,
				'reader',
				undef,
				12345.678, undef, -1, -1.25478,
				12345.678, undef, -1, -1.55748, 0.125,
				12345.678, undef, -1, -1.25748,
				12345.678, undef, -12345678910.5, -1.25748,
				12345.678, undef, -12345678910.5, -1.29748, 1.005,
				12345.678, undef, -1.23456789101, -0.005, 0.004,
				12345.678910, undef, -1.23456789101, -0.005, 0.00445,
				123456789,
				0.33333,1.66666,2.16666,3.83333,4.11111,5.22222,6.44444,7.55555,
					8.77777,9.88888,10.090909,11.18181,12.08333,13.41666,
					0.123456,0.125,0.75,(1/24),0.000005,0.000001,0.999999,
					0.99999,0.999,undef,
				0,60.99,undef,undef,1.5,55.0,59,61.345,undef,
				0.999,1.11,55.25,59.5,60.111,61.6161,undef,
				0.124,60.99,undef,undef,1.5,55.0,59,61.345,undef,
				0.999,1.11,55.25,59.5,60.111,61.6161,undef,
				0,60.99,undef,undef,1.5,55.0,59,61.345,undef,
				0.999,1.11,55.25,59.5,60.111,61.6161,undef,
				0,60.99,undef,undef,1.5,55.0,59,61.345,undef,
				0.999,1.11,55.25,59.5,60.111,61.6161,undef,
				0.124,60.99,1.5,55.0,59,61.345,0.99999,0.999999,undef,
				1900, 1904, 1970,
				10, 'My string',
			];
my			$bad_value_ref =[
				'badfile.not',
				'badfile.not',
				'badfile.not',
				'badfile.not',
				'badfile.not',
				{ string => Excel_number_0 },
			];
my			$answer_ref = [
				qr/Could not find \/ read the file: badfile\.not/,
				qr/The string -badfile.not- does not have an xlsx file extension/,
				qr/The string -badfile.not- does not have an xml file extension/,
				qr/\QThe string -badfile.not- does not match (?^:^(dom|reader|sax)$)\E/,
				12345.678, undef, -1, -1.25478,
				12346, 0, -1, -2, 0,
				12345.68, '0.00', '-1.00', -1.26,
				'12,346', 0, '-12,345,678,911', -1,
				'12,345.68', '0.00', '-12,345,678,910.50', '-1.30', '1.01',
				'1234568%','0%','-123%','-1%','0%',
				'1234567.89%','0.00%','-123.46%','-0.50%','0.45%',
				qr/\Q1.23E+\E\d*8/,
				'1/3','1 2/3','2 1/6','3 5/6','4 1/9','5 2/9','6 4/9','7 5/9','8 7/9','9 8/9',
				'10 1/11','11 2/11','12 1/12','13 5/12','10/81','1/8','3/4','1/24','1/200000',
				'1/1000000','1','1','999/1000','0/0',
				'1-1-00','3-1-00',
				qr/\Q-1900-January-0- is not a real date\E/,qr/\Q-1900-February-29- is not a real date\E/,
				'1-1-00','2-24-00','2-28-00','3-1-00',undef,
				'1-1-04','1-2-04','2-25-04','2-29-04','3-1-04','3-2-04',undef,
				'1-Jan-00','1-Mar-00',
				qr/^$/,qr/\Q-1900-February-29- is not a real date\E/,
				'1-Jan-00','24-Feb-00','28-Feb-00','1-Mar-00',undef,
				'1-Jan-04','2-Jan-04','25-Feb-04','29-Feb-04','1-Mar-04','2-Mar-04',undef,
				'1-Jan','1-Mar',
				qr/\Q-1900-January-0- is not a real date\E/,qr/\Q-1900-February-29- is not a real date\E/,
				'1-Jan','24-Feb','28-Feb','1-Mar',undef,
				'1-Jan','2-Jan','25-Feb','29-Feb','1-Mar','2-Mar',undef,
				'Jan-00','Mar-00',
				qr/\Q-1900-January-0- is not a real date\E/,qr/\Q-1900-February-29- is not a real date\E/,
				'Jan-00','Feb-00','Feb-00','Mar-00',undef,
				'Jan-04','Jan-04','Feb-04','Feb-04','Mar-04','Mar-04',undef,
				'2:58 AM','11:45 PM','12:00 PM','12:00 AM','12:00 AM',
					'8:16 AM','11:59 PM','12:00 AM',undef,
				undef, undef, qr/\QValue "1970" did not pass type constraint "EpochYear"\E/,
				1,1,0,
				qr/\Qdoes not allow key "string" to appear in hash\E/,
			];
### <where> - harder questions ...
			map{
ok			FileName->( "$test_dir$question_ref->[$_]" ),
							"Check that a good file name passes FileName: $test_dir$question_ref->[$_]";
			} ( 0..1);
			$position = 0;
dies_ok{	FileName->( $bad_value_ref->[$position] ) }
							"Check that a bad file name fails FileName: $bad_value_ref->[$position]";
like		$@, $answer_ref->[$position],
							"... and check for the correct error message";
ok			XLSXFile->( "$test_dir$question_ref->[$position]" ),
							"Check that a good file name passes XLSXFile: $test_dir$question_ref->[$position++]";
dies_ok{	XLSXFile->( $bad_value_ref->[$position] ) }
							"Check that a bad file name fails XLSXFile: $bad_value_ref->[$position]";
like		$@, $answer_ref->[$position],
							"... and check for the correct error message";
ok			XMLFile->( "$test_dir$question_ref->[$position]" ),
							"Check that a good file name passes XMLFile: $test_dir$question_ref->[$position++]";
dies_ok{	XMLFile->( $bad_value_ref->[$position] ) }
							"Check that a bad file name fails XMLFile: $bad_value_ref->[$position]";
like		$@, $answer_ref->[$position],
							"... and check for the correct error message";
ok			ParserType->( $question_ref->[$position] ),
							"Check that a good value passes ParserType: $question_ref->[$position++]";
dies_ok{	ParserType->( $bad_value_ref->[$position] ) }
							"Check that a bad value fails ParserType: $bad_value_ref->[$position]";
like		$@, $answer_ref->[$position],
							"... and check for the correct error message";
			$position++;
			my	$bad_position = $position;
			map{
is			Excel_number_0->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform Excel_ number_0 on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..3 );
			map{
is			OneFromNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform OneFromNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..4 );
			map{
is			TwoFromNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform TwoFromNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..3 );
			map{
is			ThreeFromNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform ThreeFromNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..3 );
			map{
is			FourFromNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform FourFromNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..4 );
			map{
is			NineFromNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform NineFromNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..4 );
			map{
is			TenFromNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform TenFromNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..4 );
			map{
like		ElevenFromNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform ElevenFromNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..0 );
			map{
is			TwelveFromNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform TwelveFromNum on (question place $position): " .
								($question_ref->[$position] // 'Undef') . ' -> ' . $answer_ref->[$position++];
			}( 0..23 );
#~ exit 1;
			map{
			$capture = capture_stderr{
is			FourteenFromWinExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform FourteenFromWinExcelNum on (question place $position): " . ($question_ref->[$position] // 'Undef');
			};
like		$capture, $answer_ref->[($position+2)],
							"... and check the error message";
			$position++;
			}( 0..1 );
			$position++;$position++;
			map{
is			FourteenFromWinExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform FourteenFromWinExcelNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..4 );
			map{
is			FourteenFromAppleExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform FourteenFromAppleExcelNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..6 );
			map{
			$capture = capture_stderr{
is			FifteenFromWinExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform FifteenFromWinExcelNum on (question place $position): " . ($question_ref->[$position] // 'Undef');
			};
like		$capture, $answer_ref->[($position+2)],
							"... and check the error message";
			$position++;
			}( 0..1 );
			$position++;$position++;
			map{
is			FifteenFromWinExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform FifteenFromWinExcelNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..4 );
			map{
is			FifteenFromAppleExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform FifteenFromAppleExcelNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..6 );
			map{
			$capture = capture_stderr{
is			SixteenFromWinExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform SixteenFromWinExcelNum on (question place $position): " . ($question_ref->[$position] // 'Undef');
			};
like		$capture, $answer_ref->[($position+2)],
							"... and check the error message";
			$position++;
			}( 0..1 );
			$position++;$position++;
			map{
is			SixteenFromWinExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform SixteenFromWinExcelNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..4 );
			map{
is			SixteenFromAppleExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform SixteenFromAppleExcelNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..6 );
			map{
			$capture = capture_stderr{
is			SeventeenFromWinExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform SeventeenFromWinExcelNum on (question place $position): " . ($question_ref->[$position] // 'Undef');
			};
like		$capture, $answer_ref->[($position+2)],
							"... and check the error message";
			$position++;
			}( 0..1 );
			$position++;$position++;
			map{
is			SeventeenFromWinExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform SeventeenFromWinExcelNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..4 );
			map{
is			SeventeenFromAppleExcelNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform SeventeenFromAppleExcelNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..6 );
			map{
is			EighteenFromNum->( $question_ref->[$position] ), $answer_ref->[$position],
							"Check the result of the transform EighteenFromNum on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..8 );
			map{
ok			EpochYear->( $question_ref->[$position] ),
							"Verifying the identification of valid EpochYear on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..1 );
dies_ok{	EpochYear->( $question_ref->[$position] ) }
							"Check that a bad value fails EpochYear(allowed Excel epoch start year): $question_ref->[$position]";
like		$@, $answer_ref->[$position++],
							"... and check for the correct error message";
			map{
ok			PassThroughType->( $question_ref->[$position] ),
							"Verifying PassThroughType can work on (question place $position): " . ($question_ref->[$position++] // 'Undef');
			}( 0..1 );
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