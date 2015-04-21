package Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain;
use version; our $VERSION = qv('v0.36.20');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use lib	'../../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
with	'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData';

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub get_calc_chain_position{
	my( $self, $position ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::get_calc_chain_position', );
	if( !defined $position ){
		$self->set_error( "Requested calc chain position required - none passed" );
		return undef;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Getting the calcChain position: $position" ] );
	
	# Initiate the read or reset the file if needed
	if( $self->has_position and $self->where_am_i > $position ){
		$self->start_the_file_over;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Finished resetting the file" ] );
	}
	if( !$self->has_position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the first cell" ] );
		my $found_it = $self->next_element( 'c' );
		if( $found_it < 1 ){
			$self->set_error( "No strings stored in the sharedStrings file" );
			return undef;
		}
		$self->_i_am_here( 0 );
	}
	
	while( $self->where_am_i < $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the next cell: " . ($self->where_am_i + 1) ] );
		my $result = $self->next_element( 'c' );
		if( !$result ){
			$self->_clear_location;
			$self->start_the_file_over;
			return undef;
		}
		$self->_i_am_here( $self->where_am_i + 1 );
	}
	
	my $calc_chain_node = $self->parse_element;
	$self->_i_am_here( $self->where_am_i + 1 );
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Returning shared strings node:",
	###LogSD		(($calc_chain_node) ? $calc_chain_node : '') ] );
	return $calc_chain_node;
}



#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9



#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain - A LibXML::Reader calcChain base class

=head1 SYNOPSIS

See General  L<SYNOPSIS|Spreadsheet::XLSX::Reader::LibXML::CalcChain/SYNOPSIS>
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module.  To use the general 
package for excel parsing out of the box please review the documentation for L<Workbooks
|Spreadsheet::XLSX::Reader::LibXML>, L<Worksheets
|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>.

This class is used to access the sub file calcChain.xml from an unzipped .xlsx file.  
The file to be read is generally found in the xl/ sub folder after the master file is 
unzipped.  The calcChain.xml file contains the calculation sequence and data source(s) 
used when building the value for cells with formulas.  (The formula presented in the 
L<Cell|Spreadsheet::XlSX::Reader::LibXML::Cell> instance is collected from somewhere 
else.)

This particular class uses the L<XML::LibXML::Reader> to parse the file.  The 
intent with that reader is to maintain the minimum amount of the file in working memory.  
As a consequence two things are sacrificed.  First this implementation will read the file 
serially.  This means that to read a previous element the file must start over at the 
beginning.  Second, the connection between the XMLReader and the file is broken to 
rewind the file pointer.  In an effort to minimize this pain the system file 'open' 
function is handled in a separate file handle variable so that a new system open is 
not required on each rewind.  For information regarding necessary steps to replace 
this package review the L<general|Spreadsheet::XLSX::Reader::LibXML::CalcChain/DESCRIPTION> 
documentation

This is a child class that both inherits from a parent and collates a role for full 
functional implementation.  Read the documentation for each of them as well as this 
documentation to gain a complete picture of this class.

=head2 extends

This is the parent class of this class

=head3 L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>

=head2 with

These are attached roles for additional (re-used) functionality

=head3 L<Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData>

=head2 Primary Functions

These are the primary way(s) to use this class.

=head3 get_calc_chain_position( $integer )

=over

B<Definition:> This will return the calcChain information from the identified $integer position.  
(Counting from zero).  The information is returned as a perl hash ref.

B<Accepts:> an $integer for the calcChain position

B<Returns:> a hash ref of data

=back

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

This software is copyrighted (c) 2014, 2015 by Jed Lund

=head1 DEPENDENCIES

=over

L<version>

L<perl 5.010|perl/5.10.0>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults>

L<lib>

L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>

L<Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData>

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