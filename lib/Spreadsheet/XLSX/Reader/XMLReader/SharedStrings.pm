package Spreadsheet::XLSX::Reader::XMLReader::SharedStrings;
use version; our $VERSION = version->declare("v0.1_1");

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int
    );
use lib	'../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
extends	'Spreadsheet::XLSX::Reader::XMLReader';

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9



#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _unique_count =>(
	isa			=> Int,
	writer		=> '_set_unique_count',
	reader		=> '_get_unique_count',
	clearer		=> '_clear_unique_count',
);

has _count =>(
	isa			=> Int,
	writer		=> '_set_count',
	reader		=> '_get_count',
	clearer		=> '_clear_count',
	predicate	=> '_has_count',
);
	
has +_core_element =>(
		default => 'si',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9


augment 'get_position' => sub{
	my ( $self, )	= shift;
	my	$position 	= $self->_get_requested_position;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space . '::get_position::augmented', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Reached augment::_get_position for: $position" ] );
	
	#checking if the reqested position is too far
	if( $position > $self->_get_unique_count - 1 ){
		###LogSD	$phone->talk( level => 'warn', message => [
		###LogSD		"Asking for position -$position- (from 0) but the shared string " .
		###LogSD		"max cell position is: " . ($self->_get_unique_count - 1) ] );
		return 1;#  fail
	}else{
		###LogSD	$phone->talk( level => 'debug', message =>[ "No end in sight" ] );
		return undef;#No failure
	}
};

sub _load_unique_bits{
	my( $self, $reader, $encoding ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_load_unique_bits', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the sharedStrings unique bits with reader: ", $reader,
	###LogSD			"With encoding: $encoding" ] );
	my	$node_name	= $reader->name;
	my	$found_it	= 1;
	if( $node_name ne 'sst' ){
		$found_it = $reader->nextElement( 'sst' );
	}
	if( $found_it ){
		$self->_set_unique_count( $reader->getAttribute( 'uniqueCount' ) );#
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The unique count is: " . $self->_get_unique_count ] );
		$self->_set_count( $reader->getAttribute( 'count' ) );#
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The count is: " . $self->_get_count ] );
		return undef;
	}else{
		$self->_set_error( "No 'sst' element found - can't parse this as a shared strings file" );
		$self->_clear_unique_count;
		$self->_clear_count;
		return 2;
	}
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::XMLReader::SharedStrings - sharedStrings parsing with XML::LibXML::Reader
    
=head1 DESCRIPTION

This is the XMLReader version of the Shared strings parser.  Both the XMLReader and DOM 
versions of this parser should function in the background of this package. There shouldn't 
be any need to manage this class by the user.

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader/issues>

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

B<5.010> - (L<perl>)

L<version>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<Spreadsheet::XLSX::Reader::XMLReader>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::XLSX::Reader::TempFilter>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9