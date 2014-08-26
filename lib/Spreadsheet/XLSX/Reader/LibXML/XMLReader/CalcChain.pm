package Spreadsheet::XLSX::Reader::XMLReader::CalcChain;
use version; our $VERSION = version->declare("v0.1_1");

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use lib	'../../../../../lib',;
extends	'Spreadsheet::XLSX::Reader::XMLReader';

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9
	
has +_core_element =>(
		default => 'c',
	);

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9



#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9



#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::XMLReader::CalcChain - Get a cell from the calcChain file
    
=head1 DESCRIPTION

This is mostly a stub for the XMLReader branch of the workbook level functionality.  In the 
future I would like to add the ability to know when a cell was last calculated in order to 
allow the reader to potentially re-calculate.  For now it just accesses the sheet.

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

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9