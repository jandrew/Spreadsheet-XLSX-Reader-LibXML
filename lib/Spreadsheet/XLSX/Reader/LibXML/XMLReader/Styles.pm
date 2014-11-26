package Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles;
use version; our $VERSION = qv('v0.12.4');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		InstanceOf			HashRef				Str
		Int					Bool
    );
use lib	'../../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
with	'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData',
		'Spreadsheet::XLSX::Reader::LibXML::LogSpace',
		'Spreadsheet::XLSX::Reader::LibXML::UtilFunctions',
		;

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9

my	$element_lookup ={
		numFmts			=> 'numFmt',
		fonts			=> 'font',
		borders			=> 'border',
		fills			=> 'fill',
		cellStyleXfs	=> 'xf',
		cellXfs			=> 'xf',
		cellStyles		=> 'cellStyle',
		tableStyles		=> 'tableStyle',
	};

my	$id_lookup ={
		numFmts			=> 'numFmtId',
		fonts			=> 'fontId',
		borders			=> 'borderId',
		fills			=> 'fillId',
		cellStyleXfs	=> 'xfId',
	};

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub get_format_position{
	my( $self, $position, $header ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space . '::get_format_position', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Get defined formats at position: $position",
	###LogSD			(($self->_has_sub_translation) ? '..against current stored translation: ' . $self->_get_sub_translation : undef),
	###LogSD			(($self->_has_sub_position) ? '..against current stored position: ' . $self->_get_sub_position : undef),
	###LogSD			(($header) ?  "For header: $header" : undef),
	###LogSD			(($self->_has_current_header) ? "..against stored header: " . $self->_get_current_header : undef) , ] );
	# Check header request
	if( $header and !exists( $id_lookup->{$header} ) ){
		$self->set_error( "requested header -$header- does not match the lookup list - maybe it's got a typo? ( " . join( ' - ', keys %$id_lookup ) . ')' );
	}
	
	# Check for stored final value - this only works if the target header is all that is returned
	my	$already_got_it = 0;
	if(	$header and $self->_has_current_header and
		$header eq $self->_get_current_header and
		$self->_get_sub_translation == $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Already collected this header: $header", "..and position: $position"  ] );
		$already_got_it = 1;
	}elsif( 	!$header and $self->_has_current_header and
				$self->_get_current_header eq 'cellXfs' and
				$self->_get_sub_position == $position		){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Already collected this general format at position: $position"  ] );
		$already_got_it = 1;
	}
	return $self->_get_sub_position_ref if $already_got_it;
	
	# build from scratch	
	my	$result = $self->_get_header_and_position( 'cellXfs', $position );
	return $result if ! $result;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"The ref at position -$position- is:", $result ] );
	$result = $self->_add_sub_refs( $result, $header, $position, 'cellXfs' );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"The ref at position -$position- is:", $result ] );
	return $result;
}

sub get_default_format_position{
	my( $self, $header ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space . '::get_default_format_position', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Get default format set ...", ] );
	###LogSD		$phone->talk( level => 'debug', message => [ "For header: $header" ] ) if $header;
	# Check header request
	if( $header and !exists( $id_lookup->{$header} ) ){
		$self->set_error( "requested header -$header- does not match the lookup list - maybe it's got a typo? ( " . join( ' - ', keys %$id_lookup ) . ')' );
	}
	
	# Get base ref
	my	$result = $self->_get_header_and_position( 'cellStyleXfs', 0 );
	return $result if ! $result;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Default formats:", $result ] );
	$result = $self->_add_sub_refs( $result, $header, 0, 'cellStyleXfs' );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final formats:", $result ] );
	return $result;
}

sub get_sub_format_position{
	my( $self, $position, $header ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space . '::get_sub_format_position', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Get sub format for -$header- at position: $position",
	###LogSD			(($self->_has_sub_position) ? '..against current stored position: ' . $self->_get_sub_position : undef),
	###LogSD			(($header) ?  "For header: $header" : undef),
	###LogSD			(($self->_has_current_header) ? "..against stored header: " . $self->_get_current_header : undef) , ] );
	
	# Check the validaty of the request
	my	$has_required = 1;
	if( !defined $header ){
		$has_required = 0;
		$self->set_error( '$header is a required value for the method - get_sub_format_position( $position, $header )'  );
	}elsif( !defined $position ){
		$has_required = 0;
		$self->set_error( '$position is a required value for the method - get_sub_format_position( $position, $header )'  );
	}elsif( !exists( $id_lookup->{$header} ) ){
		$has_required = 0;
		$self->set_error( "requested header -$header- does not match the lookup list - maybe it's got a typo? ( " . join( ' - ', keys %$id_lookup ) . ')' );
	}
	return undef if !$has_required;
	
	# Check for stored final value - this only works if the target header is all that is returned
	if(	$self->_has_current_header and
		$header eq $self->_get_current_header and
		$self->_get_sub_position == $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Already collected this header: $header", "..and position: $position"  ] );
		return $self->_get_sub_position_ref;
	}
	
	# build from scratch
	my	$result = $self->_get_header_and_position( $header, $position );
	return $result if ! $result;
	$self->_set_current_header( $header );
	$self->_set_sub_position_ref( { $header => $result } );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"The ref at position -$position- is:", $result ] );
	return { $header => $result };
	
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _style_block_size =>(
		isa		=> HashRef,
		traits	=> ['Hash'],
		writer	=> '_set_style_block_size',
		handles	=>{
			_get_block_size => 'get',
		},
	);

has _current_header =>(
		isa		=> Str,
		writer	=> '_set_current_header',
		reader	=> '_get_current_header',
		clearer	=> '_clear_current_header',
		predicate	=> '_has_current_header',
		trigger	=> \&_clear_sub_translation,
	);

has _sub_translation =>(
		isa		=> Int,
		writer	=> '_set_sub_translation',
		reader	=> '_get_sub_translation',
		clearer	=> '_clear_sub_translation',
		predicate	=> '_has_sub_translation',
		trigger	=> \&_clear_sub_position_ref,
	);

has _sub_position =>(
		isa		=> Int,
		writer	=> '_set_sub_position',
		reader	=> '_get_sub_position',
		clearer	=> '_clear_sub_position',
		predicate	=> '_has_sub_position',
		trigger	=> \&_clear_current_header,
	);

has _sub_position_ref =>(
		isa		=> HashRef,
		writer	=> '_set_sub_position_ref',
		reader	=> '_get_sub_position_ref',
		clearer	=> '_clear_sub_position_ref',
		predicate	=> '_has_sub_position_ref',
	);
	
has _last_recorded =>(
		isa		=> Bool,
		writer	=> '_set_last_recorded',
		reader	=> '_get_last_recorded',
		default	=> 0,
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _load_unique_bits{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_load_unique_bits', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Loading the counts and attributes of all the style types",
	###LogSD			'bytes consumed: ' . $self->byte_consumed, 'At node: ' . $self->node_name ] );
	if( $self->node_name ne 'styleSheet' ){
		$self->next_element( 'styleSheet' );
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		'bytes consumed: ' . $self->byte_consumed, 'At node: ' . $self->node_name ] );
	}
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		'lower level ? bytes consumed: ' . $self->byte_consumed, 'At node: ' . $self->node_name ] );
	my	$top_level_ref = $self->parse_element( 2 );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Resulting parse:", $top_level_ref ] );
	$self->_set_style_block_size( $top_level_ref );
	$self->start_the_file_over;
	$self->next_element( 'numFmts' );
	my	$number_ref = $self->parse_element;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Number format list:", $number_ref ] );
	my	$translations = $self->get_defined_excel_format_list;
	for my $format ( @{$number_ref->{list}} ){
		$translations->[$format->{numFmtId}] = "$format->{formatCode}";
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		'loaded format: ' . $translations->[$format->{numFmtId}] ] );
	}
	$self->set_defined_excel_format_list( $translations );
}

sub _get_header_and_position{
	my( $self, $target_header, $target_position ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_get_header_and_position', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"getting the ref for target header: $target_header",
	###LogSD			"..and position: $target_position"						] );
	if( $target_header eq 'numFmts' ){
		my $format_string = $self->get_defined_excel_format( $target_position );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Special numFmts call: $target_position", "..returned: " . 
		###LogSD		(($format_string) ? $format_string : ''),	] );
		my $conversion = $self->parse_excel_format_string( $format_string, "Excel__$target_position" );
		return $conversion;
	}
		
	my $test_ref = $self->_get_block_size( $target_header );
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Recorded size of -$target_header- is:", $test_ref ] );
	if( !$test_ref ){
		$self->set_error( "Target header -$target_header- not found in the loaded Styles sheet" );
		return undef;
	}elsif( $test_ref->{count} < $target_position + 1 ){
		$self->set_error( "Header -$target_header- does not extend to position: $target_position" );
		return undef;
	}
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"The target data should exist" 			] );
	if( $self->_has_current_header ){
		if(	$self->_get_current_header ne $target_header or
			$target_header eq 'cellXfs' or
			$self->_get_sub_position > $target_position		){
			###LogSD		$phone->talk( level => 'debug', message => [
			###LogSD			"a bridge to far - reset" 			] );
			$self->_set_last_recorded( 0 );
			$self->_clear_current_header;
			$self->_clear_sub_position;
			$self->_clear_sub_position_ref;
			$self->_clear_sub_translation;
			$self->start_the_file_over;
		}
	}
	
	my	$element_name = $element_lookup->{$target_header};
	my ( $sub_position, $last_recorded, );# $changed_position
	###LogSD		$phone->talk( level => 'trace', message => [
	###LogSD			"Getting to the right header" ] );
	if( $self->_has_current_header and $self->_get_current_header eq $target_header ){
		$sub_position = $self->_get_sub_position;
		###LogSD		$phone->talk( level => 'trace', message => [
		###LogSD			"Already at: $target_header", "..and sub position: $sub_position" ] );
	}else{
		my	$result = $self->next_element( $target_header );
		if( !$result ){
			###LogSD		$phone->talk( level => 'trace', message => [
			###LogSD			"Failed to find the node resetting the file" ] );
			$self->start_the_file_over;
			$result = $self->next_element( $target_header );
			$self->_set_last_recorded( 0 );
		}
		###LogSD		$phone->talk( level => 'trace', message => [
		###LogSD			"Arrived at: " . $self->node_name,
		###LogSD			"Result of advancing to -$target_header- : $result" ] );
		$result = $self->next_element( $element_name );
		###LogSD		$phone->talk( level => 'debug', message => [
		###LogSD			"Result of advancing to the first element -$element_name- is: $result",
		###LogSD			'..at node: ' . $self->node_name, '..and byte position: ' . $self->byte_consumed ] );
		$sub_position = 0;
		#~ $changed_position = 1;
	}
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Getting to the right element position from sub position: $sub_position" ] );
	if( $target_position > $sub_position ){
		$sub_position++ if $self->_get_last_recorded == 1;
		$self->_set_last_recorded( 0 );
		#~ my $ref = $self->parse_element;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Sub position now: $sub_position",] );# $ref 
		#~ $changed_position = 1;
		while( $target_position > $sub_position ){
			my $result = $self->next_element( $element_name );
			$sub_position++;
			###LogSD		$phone->talk( level => 'debug', message => [
			###LogSD			"Result of advancing to the next element -$element_name- is: $result",
			###LogSD			'..at node: ' . $self->node_name,
			###LogSD			'..and byte position: ' . $self->byte_consumed, "..for positon: $sub_position" ] );
		}
	}
	$self->_set_sub_position( $sub_position );
	###LogSD		$phone->talk( level => 'trace', message => [
	###LogSD			"pulling the element for the position: $sub_position" ] );
	my	$position_definition = $self->parse_element;#( 5 );
	$self->_set_last_recorded( 1 ) if $self->node_name eq $element_name;
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Returning position ref:", $position_definition ] );
	return $position_definition;
}

sub _add_sub_refs{
	my( $self, $base_ref, $header, $super_position, $base_header ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_sub_refs', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building the sub ref for:", $base_ref,
	###LogSD			(($header) ? "..focused on header: $header" : undef), ] );
	if( $header ){	
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The call is to add data for just one specific header: $header", ] );
		if( exists( $base_ref->{$id_lookup->{$header}} ) ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Collecting subdata for position: $base_ref->{$id_lookup->{$header}}", ] );
			my	$sub_result = $self->_get_header_and_position( $header, $base_ref->{$id_lookup->{$header}} );
			return undef if !$sub_result;
			$self->_set_current_header( $header );
			$self->_set_sub_translation( $super_position );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'Setting subref:', $sub_result, ] );
			$self->_set_sub_position_ref( { $header => $sub_result } );
			$base_ref = { $header => $sub_result };
		}else{
			$self->set_error( "requested header -$header- generally exists but has no pointer to a sub definition in the Styles file." );
			$self->_set_sub_position( 0 );
			$self->_clear_sub_position;
			return undef;
		}
	}else{		
		for my $header ( keys %$id_lookup ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Checking for subdata of header: $header" ] );
			if( exists( $base_ref->{$id_lookup->{$header}} ) ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Retreiving the data at subposition: $base_ref->{$id_lookup->{$header}}" ] );
				$base_ref->{$header} = $self->_get_header_and_position( $header, $base_ref->{$id_lookup->{$header}} );
			}
		}
		$self->_set_sub_position( $super_position );
		$self->_set_current_header( $base_header );
		$self->_set_sub_position_ref( $base_ref);
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"The final ref is:", $base_ref ] );
	return $base_ref;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles - LibXML::Reader parser of Styles
    
=head1 DESCRIPTION

POD not written yet!

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Nothing L<yet|/SUPPORT>

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

L<Spreadsheet::XLSX::Reader::LibXML>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::ParseExcel> - Excel 2003 and earlier

L<Spreadsheet::XLSX> - 2007+

L<Spreadsheet::ParseXLSX> - 2007+

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=over

All lines in this package that use Log::Shiras are commented out

=back

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9