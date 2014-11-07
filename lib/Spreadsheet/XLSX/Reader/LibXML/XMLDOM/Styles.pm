package Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles;
use version; our $VERSION = qv('v0.5_1');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Type::Utils -all;
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
use Type::Coercion;
use DateTimeX::Format::Excel v0.12;
use lib	'../../../../../../lib',;
with 'Spreadsheet::XLSX::Reader::LibXML::LogSpace';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML::Types v0.1 qw(
		XMLFile
		PassThroughType
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
with 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData';

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9



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
		reader	=> 'get_excel_region',
	);
	
has defined_excel_translations =>(
		isa		=> ArrayRef[Maybe[HashRef[Maybe[HashRef]]]],
		traits	=> ['Array'],
		default	=> sub{ [
				{ 	'1900' =>{ en => PassThroughType->plus_coercions( ZeroFromNum )}},
				{ 	'1900' =>{ en => OneFromNum }},
				{ 	'1900' =>{ en => TwoFromNum }},
				{ 	'1900' =>{ en => ThreeFromNum }},
				{ 	'1900' =>{ en => FourFromNum }},
				undef,
				undef,
				undef,
				undef,
				{ 	'1900' =>{ en => NineFromNum }},
				{ 	'1900' =>{ en => TenFromNum }},
				{ 	'1900' =>{ en => ElevenFromNum }},
				{ 	'1900' =>{ en => TwelveFromNum }},
				{ 	'1900' =>{ en => TwelveFromNum }},
				{ 	'1900' =>{ en => FourteenFromWinExcelNum },
					'1904' =>{ en => FourteenFromAppleExcelNum }},
				{	'1900' =>{ en => FifteenFromWinExcelNum },
					'1904' =>{ en => FifteenFromAppleExcelNum }},
				{	'1900' =>{ en => SixteenFromWinExcelNum },
					'1904' =>{ en => SixteenFromAppleExcelNum }},
				{	'1900' =>{ en => SeventeenFromWinExcelNum },
					'1904' =>{ en => SeventeenFromAppleExcelNum }},
				{	'1900' =>{ en => EighteenFromNum }},
		] },
		handles =>{
			_get_translation_array_position => 'get',
			_set_translation_array_position => 'set',
		},
		reader	=> '_get_translation_array',
		writer	=> '_set_translation_array',
	);
	
has	epoch_year =>(
		isa		=> Int,
		reader	=> 'get_epoch_year',
		writer	=> 'set_epoch_year',
		default	=> 1900,
	);

has	error_inst =>(
		isa		=> InstanceOf[ 'Spreadsheet::XLSX::Reader::LibXML::Error' ],
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
		clearer	=> '_clear_error_inst',
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub parse_excel_format_string{# Currently only handles dates and times
	my( $self, $format_string, $ID ) = @_;
	no warnings 'uninitialized';
	my	$target_name = "DateTime${ID}FromNum";
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space .  '::parse_excel_format_string', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"parsing the custom excel format string: $format_string",
	###LogSD			"To a conversion named: $target_name", ] );
	#~ my ( @coercion_list );
	my ( $coercion, @coercion_list );
	for my $format_sub_string ( split /;/, $format_string ){
		$format_sub_string =~ /([dymh]*)?(@)?/;
		###LogSD	no warnings 'uninitialized';
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD			"Testing substring: $format_sub_string"  ] );
		#~ my ( $type, $coercion );
		if( $format_sub_string =~ /[dymh]/ ){
			#~ $type = Num;
			push @coercion_list, $self->_build_date_time_coercion( $format_sub_string, $ID );
		}elsif( $format_sub_string =~ /@/  ){
			next;
			push @coercion_list, Str, sub{ $_ };
		}else{
			$self->set_error( "Can't parse ID number -$ID- format sub-string |$format_sub_string| yet" );
		}
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD			"Current coercion list:", @coercion_list ] );
	}
	$coercion = Type::Coercion->new(
	   name              => $target_name,
	   #~ type_constraint   => Str,
	   type_coercion_map => [ @coercion_list ],
	);
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"string parsed to:", $coercion] );
	return $coercion;#$coercion_list[1]
}

sub get_format_position{
	my( $self, $position, $header ) = @_;
	my	$epoch_year	= $self->get_epoch_year;
	my	$region		= $self->get_excel_region;
	my	$conversion;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space . '::get_format_position', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"For epoch: $epoch_year",  "and region: $region",
	###LogSD			"Get predefined conversion at position: $position", ] );
	###LogSD		$phone->talk( level => 'debug', message => [ "For header: $header" ] ) if $header;
	$conversion = $self->get_cellXfs;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"All conversions:", $conversion ] );
	$conversion = $conversion->[$position];
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"intermediate conversion:", $conversion ] );
	$conversion = $self->_extract_header( $conversion, $header ) if $header;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final conversion:", $conversion ] );
	return $conversion;
}

sub get_default_format_position{
	my( $self, $header ) = @_;
	my	$epoch_year	= $self->get_epoch_year;
	my	$region		= $self->get_excel_region;
	my	$conversion;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space . '::get_default_format_position', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"For epoch: $epoch_year",  "and region: $region",
	###LogSD			"Get default format set ...", ] );
	###LogSD		$phone->talk( level => 'debug', message => [ "For header: $header" ] ) if $header;
	$conversion = $self->get_cellStyleXfs->[0];
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Default conversion:", $conversion ] );
	$conversion = $self->_extract_header( $conversion, $header ) if $header;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final conversion:", $conversion ] );
	return $conversion;
}
	

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _style_in_perl =>(
	isa		=> HashRef,
	traits	=> ['Hash'],
	writer	=> '_set_style_in_perl',
	handles	=>{
		get_numFmts			=>[ get => 'numFmts' ],
		get_cellXfs			=>[ get => 'cellXfs' ],
		get_fonts			=>[ get => 'fonts' ],
		get_numFmts			=>[ get => 'numFmts' ],
		get_cellStyleXfs	=>[ get => 'cellStyleXfs' ],
		get_tableStyles		=>[ get => 'tableStyles' ],
		get_fills			=>[ get => 'fills' ],
		get_borders			=>[ get => 'borders' ],
		get_dxfs			=>[ get => 'dxfs' ],
		get_cellStyles		=>[ get => 'cellStyles' ],
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
	###LogSD	$phone->talk( level => 'debug', message => [ "Good reader built" ] );
	
	# Get file encoding
	my	$encoding	= $reader->encoding();
	$self->_set_encoding( $encoding );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"Encoding of file is: $encoding" ], );
	
	# Process the style sheet
	my ( $list_name, $list_ref ) =	$self->process_element_to_perl_data(
										$reader->getElementsByTagName( 'styleSheet' )
									);
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"intermediate styles data ref:", $list_ref, ], );
	$self->_set_style_in_perl( $list_ref );
	
	# Read custom numFmt's into a hash
	my	$x = 0;
	for my $format_node ( @{$list_ref->{numFmts}} ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD			"Custom number format node -$x- is: ", $format_node ], );
		my $ID = $format_node->{numFmtId};
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD			"Custom number format node -$x- for position -$ID- is: $format_node->{formatCode}" ], );
		$self->_set_translation_array_position( $ID, {
			$self->get_epoch_year =>{
				$self->get_excel_region => $self->parse_excel_format_string(
					$format_node->{formatCode}, $ID,
				)
			}, }
		);
		$x++;
	}
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD			"Final custom number formats: ", $self->_get_translation_array], );
	my @position_list;
	for my $pos ( 0 .. $#{$list_ref->{cellXfs}} ){
		( $list_ref->{cellXfs}->[$pos], my $error_id ) =
			$self->_load_data_to_format( $list_ref->{cellXfs}->[$pos] );
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD			"Final position is:", $list_ref->{cellXfs}->[$pos] ], );
		push @position_list, $error_id if defined $error_id;
	}
	if( scalar( @position_list ) > 0 ){
		$self->set_error( "Missing number formats for defined excel translations at positions( " .
			join( ' - ', @position_list ) . ' )' );
	}
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD			"intermediate styles data ref:", $list_ref, ], );
	for my $pos ( 0 .. $#{$list_ref->{cellStyleXfs}} ){
		( $list_ref->{cellStyleXfs}->[$pos], my $error_id ) =
			$self->_load_data_to_format( $list_ref->{cellStyleXfs}->[$pos] );
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD			"Final position is:", $list_ref->{cellStyleXfs}->[$pos] ], );
		push @position_list, $error_id if defined $error_id;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD			"Final styles data ref:", $list_ref, ], );
	$self->_set_style_in_perl( $list_ref );
	
	#clear loaders
	$reader = undef;
	return 1;
}

sub _build_date_time_coercion{
	my( $self, $format_string, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space .
	###LogSD			'::parse_excel_format_string::_build_date_time_coercion', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"parsing the custom excel date time string: $format_string",
	###LogSD			"For epoch year: " . $self->get_epoch_year, ] );
	$format_string =~ /(\[[^\]]+\])?(.+)/;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"No differentiation available for: $1" ] ) if $1;
	my	$excel_date_string = $2;
	if( $excel_date_string =~ /[dy]/ ){
		$excel_date_string =~ s/m/M/g;
	}
	my	@args_list = ( $self->get_epoch_year == 1904 ) ? ( system_type => 'apple_excel' ) : ();
	my	$converter = DateTimeX::Format::Excel->new( @args_list );
	return ( Str, sub{ 
					my	$num = $_[0];
					my	$dt = $converter->parse_datetime( $num );
					return $dt->format_cldr( $excel_date_string );
				} );
}

sub _load_data_to_format{
	my( $self, $element, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space . '::_load_data_to_format', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"parsing the format element:", $element ] );
	my $error_id;
	if( exists $element->{numFmtId} ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD			"There should be a number format at: $element->{numFmtId}" ], );
		my	$format_ref = $self->_get_translation_array_position( $element->{numFmtId} );
		###LogSD	no warnings 'uninitialized';
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD			"The number conversion is:", $format_ref ], );
		if( $format_ref ){
			$format_ref = ( exists $format_ref->{$self->get_epoch_year} ) ?
				$format_ref->{$self->get_epoch_year} : $format_ref->{'1900'} ;
			$format_ref = ( exists $format_ref->{$self->get_excel_region} ) ?
				$format_ref->{$self->get_excel_region} : $format_ref->{en} ;
		}elsif( $element->{numFmtID} ){
			$error_id = $element->{numFmtID};
			$format_ref = ZeroFromNum;
		}else{
			$format_ref = ZeroFromNum;
		}
		$element->{NumberFormat} = $format_ref;
	}
	if( exists $element->{fontId} ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD			"There should be a font format at: $element->{fontId}" ], );
		$element->{font} = $self->get_fonts->[$element->{fontId}];
	}
	if( exists $element->{fillId} ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD			"There should be a fill format at: $element->{fillId}" ], );
		$element->{fill} = $self->get_fills->[$element->{fillId}];
	}
	if( exists $element->{borderId} ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD			"There should be a border format at: $element->{borderId}" ], );
		$element->{border} = $self->get_borders->[$element->{borderId}];
	}
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD			"Final position is:", $element ], );
	return( $element, $error_id );
}

sub _extract_header{
	my( $self, $element, $header ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space . '::_extract_header', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Looking for header -$header- in the element:", $element ] );
	if( $header ){
		if( exists $element->{$header} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Found header -$header- in the element" ] );
			$element = $element->{$header};
		}else{
			$self->set_error( "Header -$header- is not an available option in the format list" );
		}
	}
	return $element;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles - LibXML DOM parser of Styles
    
=head1 DESCRIPTION

This is the module that is used to apply any style definitions listed in the sheet.  See 
L<Spreadsheet::XLSX::Reader::LibXML::Worksheet> for a way to apply other styles to the 
output.  The current styles coverage is minimal and will expand over time.  In general if 
I didn't write the excel version of a style implementation this module will use the 
pass-through style.

=head1 SUPPORT

=over

L<github Spreadsheet-XLSX-Reader-LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Add some L<Data::Walk::Graft> magic to the defined_excel_translations attribute so 
this list can be managed by detail.

There are a lot of features still to be added. This module is very much a work in progress.

=back

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