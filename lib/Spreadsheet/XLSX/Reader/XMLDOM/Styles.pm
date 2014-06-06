package Spreadsheet::XLSX::Reader::XMLDOM::Styles;
use version; our $VERSION = version->declare("v0.1_1");

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int
		Str
		Maybe
		Num
		HashRef
		ArrayRef
		CodeRef
		Object
		ConsumerOf
		InstanceOf
    );
use XML::LibXML;
my	$chunk_parser = XML::LibXML->new;
use XML::LibXML::Reader;
use Type::Coercion;
use DateTimeX::Format::Excel;
use lib	'../../../../../lib',;
with 'Spreadsheet::XLSX::Reader::LogSpace';
###LogSD	use Log::Shiras::Telephone;# Fix with CPAN release of Log::Shiras
use Spreadsheet::XLSX::Reader::Types v0.1 qw(
		XMLFile
		NumberFormat
		NumberFormats
		ZeroFromNum
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

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9

my	$win_excel_converter	= DateTimeX::Format::Excel->new;
my	$apple_excel_converter	= DateTimeX::Format::Excel->new( system_type => 'apple_excel' );


#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has file_name =>(
	isa			=> XMLFile,
	predicate	=> 'has_file_name',
	trigger		=> \&_parse_the_file,
	required	=> 1,
);
	
has excel_region =>(
		isa		=> Str,
		default	=> 'en',
	);
	
has defined_excel_translations =>(
		isa		=> HashRef[HashRef[ArrayRef]],
		traits	=> ['Hash'],
		default	=> sub{
			{
				en =>{
					'1900' =>[
						ZeroFromNum,
						OneFromNum,
						TwoFromNum,
						ThreeFromNum,
						FourFromNum,
						undef, undef, undef, undef,
						NineFromNum,
						TenFromNum,
						ElevenFromNum,
						TwelveFromNum,
						TwelveFromNum,
						FourteenFromWinExcelNum,
						FifteenFromWinExcelNum,
						SixteenFromWinExcelNum,
						SeventeenFromWinExcelNum,
						EighteenFromNum,
					],
					'1904' =>[
						undef, undef, undef, undef, undef,
						undef, undef, undef, undef, undef,
						undef, undef, undef, undef,
						FourteenFromAppleExcelNum,
						FifteenFromAppleExcelNum,
						SixteenFromAppleExcelNum,
						SeventeenFromAppleExcelNum,
						undef,
					],
				},
			}
		},
		handles =>{
			_get_region_excel_formats => 'get',
			_has_region_excel_formats => 'exists',
		},
	);
	
has	epoch_year =>(
		isa		=> Int,
		reader	=> 'get_epoch_year',
		writer	=> 'set_epoch_year',
		required => 1,
	);

has	error_inst =>(
		isa		=> InstanceOf[ 'Spreadsheet::XLSX::Reader::Error' ],
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
		clearer	=> '_clear_error_inst',
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9



#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _number_formats =>(
	isa		=> NumberFormats,
	traits	=> ['Array'],
	writer	=> '_set_number_formats',
	handles	=>{
		get_number_format => 'get',
	},
);

has _default_number_format =>(
	isa		=> NumberFormat,
	writer	=> '_set_default_number_format',
	reader	=> 'get_default_number_format',
);

has _font_formats =>(
	isa		=> ArrayRef[InstanceOf['XML::LibXML::Element']],
	traits	=> ['Array'],
	writer	=> '_set_font_list',
	handles	=>{
		get_font_definition => 'get',
	},
);

has _file_encoding =>(
		isa		=> Str,
		writer	=> '_set_encoding',
		reader	=> 'encoding',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _parse_the_file{
	my( $self, $new_file, $old_file ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space .  '::_parse_the_file', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Loading the DOM reader with file: $new_file" ] );
	
	# Set the reader file
	my	$reader = XML::LibXML->load_xml( location => $new_file );
	return undef if !$reader;
	#~ $self->_set_xml_parser( $reader );
	###LogSD	$phone->talk( level => 'debug', message => [ "Good reader built" ] );
	
	# Get file encoding
	my	$encoding	= $reader->encoding();
	$self->_set_encoding( $encoding );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"Encoding of file is: $encoding" ], );
	
	my	$region		= $self->excel_region;
	my	$epoch_year = $self->get_epoch_year;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"Working with: $region",
	###LogSD			"And epoch_year: $epoch_year" ], );
	
	# Read custom numFmt's into a hash
	my $number_format_hash;
	my @numFmt_nodes = $reader->getElementsByTagName('numFmt');#getElementsByTagName( 'numFmt' );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"numFmt nodes: ", @numFmt_nodes ], );
	my	$x = 0;
	for my $format_node ( @numFmt_nodes ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD			"Custom number format node -$x- is: " . $format_node ], );
		my $ID = $format_node->getAttribute( 'numFmtId' );
		$number_format_hash->{$ID} = $self->_parse_excel_format_string(
			$format_node->getAttribute( 'formatCode' ), $ID, $epoch_year,
		);
		$x++;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"Final custom number formats: ", $number_format_hash ], );
	
	# Get all font settings
	my @font_nodes = $reader->getElementsByTagName( 'font' );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"font nodes: ", @font_nodes ], );
	
	# Build the information for each format type
	my ( $general_cellXfs ) = $reader->getElementsByTagName( 'cellStyleXfs' );
	my	$general_style = ( $general_cellXfs->getChildrenByTagName( 'xf' ) )[0];
	my ( $cellXfs_nodes ) = $reader->getElementsByTagName( 'cellXfs' );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"cellXfs node: ", $cellXfs_nodes ], );
	$x = 0;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"Trying to resolve the number format postitions for region: $region",
	###LogSD			"And epoch_year: $epoch_year" ], );
		my	$number_format_list;
	for my $format_node ( $general_style, $cellXfs_nodes->getChildrenByTagName( 'xf' ) ){
		my	$translation_ref;
		my	$test_epoch	= $epoch_year;
		my	$test_region = $region;
		my	$number_ref;
		@$number_ref{
			'borderId', 'fillId', 'font', 'xfId', 'applyNumberFormat',
			'pivotButton', 'applyAlignment', 'applyFont' 
		} = (
			$format_node->getAttribute( 'borderId' ),
			$format_node->getAttribute( 'fillId' ),
			$font_nodes[$format_node->getAttribute( 'fontId' )],
			$format_node->getAttribute( 'xfId' ),
			$format_node->getAttribute( 'applyNumberFormat' ),
			$format_node->getAttribute( 'pivotButton' ),
			$format_node->getAttribute( 'applyAlignment' ),
			$format_node->getAttribute( 'applyFont' )
		);
		if( $number_ref->{applyAlignment} ){
			my( $alignment_node ) = $format_node->getChildrenByTagName( 'alignment' );
			$number_ref->{alignment}->{horizontal} = $alignment_node->getAttribute( 'horizontal' );
			$number_ref->{alignment}->{vertical} = $alignment_node->getAttribute( 'vertical' );
		}
		my	$number_ID	= $format_node->getAttribute( 'numFmtId' );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD			"Parsing number format position -$x- for value: $number_ID",
		###LogSD			"For region: $region", $number_ref ], );
		my	$found_answer = 0;
		my	$max_tests = 10;
		if( exists $number_format_hash->{$number_ID} ){
			$found_answer = 1;
			$number_ref->{translation}	=	$number_format_hash->{$number_ID};
			$number_format_list->[$x]	=	$number_ref;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD			"Using customer number format -$number_ID-:",
			###LogSD			$number_format_hash->{$number_ID}  			], );
		}
		while( !$found_answer and $max_tests > 0 ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD			"Looking for number format -$number_ID- for region " .
			###LogSD			"-$test_region- and epoch -$test_epoch-:" ], );
			my	$region_ref;
			if( $self->_has_region_excel_formats( $test_region ) ){
				$region_ref = $self->_get_region_excel_formats( $region );
			}else{
				$test_region = 'en';
			}
			if( $region_ref ){
				if( exists $region_ref->{$test_epoch} ){
					if(	( $test_region eq 'en' and $test_epoch == 1900 ) or
						defined $region_ref->{$test_epoch}->[$number_ID] ){
						$number_ref->{translation}	=	$region_ref->{$test_epoch}->[$number_ID];
						$number_format_list->[$x]	=	$number_ref;
						$found_answer = 1;
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD			"Using customer number format -$number_ID- for region " .
						###LogSD			"-$test_region- and epoch -$test_epoch-:",
						###LogSD			$region_ref->{$test_epoch}->[$number_ID] 			], );
					}else{
						$test_epoch = 1900;
					}
				}else{
					$test_epoch = 1900;
				}
			}
		}
		$number_ref->{translation} //= $self->_get_region_excel_formats( 'en' )->{1900}->[0];
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD			"Number node -$x- is: ", $number_format_list->[$x],], );
		$x++;
	}
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD			"Final Number formats: ", $number_format_list,
	###LogSD			"Font Descriptions are:", @font_nodes,
	###LogSD			"NumberFormat error:", NumberFormats->validate_explain( $number_format_list ),  ], ) ;
	my $default = shift @$number_format_list;
	NumberFormat->( $default );
	$self->_set_default_number_format( $default );
	$self->_set_number_formats( $number_format_list );
	$self->_set_font_list( \@font_nodes );
	$self->_clear_error_inst;
	return 1;
	
}

#~ sub _set_types_name_space{
	#~ my( $self, $name_space, ) = @_;
	#~ ###LogSD	my	$phone = Log::Shiras::Telephone->new(
	#~ ###LogSD					name_space 	=> $name_space .  '::_set_types_name_space', );
	#~ ###LogSD		$phone->talk( level => 'debug', message => [
	#~ ###LogSD			"Setting the types name_space to: $name_space", ] );
	#~ no	warnings 'once';
	#~ $Spreadsheet::XLSX::Reader::Types::name_space = $name_space;
	#~ use	warnings 'once';
	#~ return 1;
#~ }

sub _parse_excel_format_string{# Currently only handles dates and times
	my( $self, $format_string, $ID, $epoch_year ) = @_;
	my	$phone;
	my	$target_name = "C${ID}FromNum";
	###LogSD	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space .  '::_parse_excel_format_string', );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"parsing the custom excel format string: $format_string",
	###LogSD			"For epoch year: $epoch_year",
	###LogSD			"To a conversion named: $target_name", ] );
	my @list = $format_string =~ /(\[[^\]]+\])([^;]+);?([^;]*);?([^;]*);?([^;]*)/;
	while( !$list[-1] ){ pop @list };#clear blanks
	pop @list if $list[-1] =~/@/;#Ignore text
	shift @list if $list[0] =~/\[/;
	if( scalar( @list ) > 1 ){
		$self->set_error( "Can't handle number ranges yet: " . join( ' - ', @list) );
		return ZeroFromNum;
	}elsif( $list[0] !~ /[dymh]/ ){
		$self->set_error( "Only custom date formats are supported so far: " . join( ' - ', @list ) );
		return ZeroFromNum;
	}
	if( $list[0] =~ /[dy]/ ){
		$list[0] =~ s/m/M/g;
	}
	my	$converter = ( $epoch_year == 1904 ) ? $apple_excel_converter : $win_excel_converter;
	my	$conversion = Type::Coercion->new(
			name	=> $target_name,
			to_type => InstanceOf[ 'DateTime' ],
			from	=> Maybe[Num],
			via		=> sub{
							my $num = $_;
							return undef if( !defined $num );
							$converter->parse_datetime( $num )->format_cldr( $list[0] );
						},
		);
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"string parsed to:",  @list, "Returning:", $conversion ] );
	return $conversion;
}
	

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable(
	inline_constructor => 0,
);
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::DOM::Styles - LibXML DOM parser of Styles
    
=head1 DESCRIPTION

This is the module that is used to apply any style definitions listed in the sheet.  See 
L<Spreadsheet::XLSX::Reader::Worksheet> for a way to apply other styles to the output.  
The current styles coverage is minimal and will expand over time.  In general if I didn't 
write the excel version of a style implementation this module will use the pass-through style.

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader/issues>

=back

=head1 TODO

There are a lot of features still to be added. This module is very much a work in progress.

=over

=item B<1.> implement more of the L<standard number formats
|http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat(v=office.14).aspx>

=back

=head1 AUTHOR

=over

=item Jed Lund

=item jandrew@cpan.org

=back

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

This software is copyrighted (c) 2014 by Jed Lund

=head1 DEPENDENCIES

=over

B<5.010> - (L<perl>)

L<version>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<XML::LibXML>

L<XML::LibXML::Reader>

L<Type::Coercion>

L<DateTimeX::Format::Excel>

L<Spreadsheet::XLSX::Reader::LogSpace>

L<Spreadsheet::XLSX::Reader::Types>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::XLSX::Reader::TempFilter>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9