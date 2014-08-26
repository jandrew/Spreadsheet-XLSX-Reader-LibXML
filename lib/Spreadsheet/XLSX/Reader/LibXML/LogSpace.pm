package Spreadsheet::XLSX::Reader::LibXML::LogSpace;
use version; our $VERSION = qv('v0.4.2');

use Moose::Role;
use Types::Standard qw(
		Str
    );
use lib	'../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has log_space =>(
		is		=> 'ro',
		isa		=> Str,
		reader	=> 'get_log_space',
		writer	=> 'set_log_space',
		default	=> __PACKAGE__,
		trigger	=> \&_set_types_log_space,
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _set_types_log_space{
	my( $self, $log_space, ) = @_;
	$log_space .= '::Types';
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $log_space .  '::_set_types_log_space', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the types name_space to: $log_space", ] );
	no	warnings 'once';
	$Spreadsheet::XLSX::Reader::LibXML::Types::log_space = $log_space;
	use	warnings 'once';
	return 1;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;

1;
# The preceding line will help the module return a true value

#########1 main POD docs      3#########4#########5#########6#########7#########8#########9

__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::LogSpace - Role to manage logging name space

=head1 DESCRIPTION

Normally the attribute justs belong in the package but it is nice to have in a 
pluggable role for sub unit testing.

=head1 SYNOPSIS
	
	#!perl
	package MyPackage;
	with 'Spreadsheet::XLSX::Reader::LibXML::LogSpace';

=head2 Attributes

Data passed to new when creating an instance of the consuming class.  For modification of 
these attributes see the listed L</Methods>.

=head3 log_space

=over

B<Definition:> This is provided for external use by the logging package L<Log::Shiras
|https://github.com/jandrew/Log-Shiras>.

B<Default> __PACKAGE__

B<Range> Any string, but Log::Shiras will look for '::' separators
		
=back

=head2 Methods

This is a method to access the attribute.

=head3 get_log_space

=over

B<Definition:> This is the way to read the set name_space. (there is no way to modify it)

B<Accepts:>Nothing

B<Returns:> the 'name_space' value

=back

=head1 SUPPORT

=over

L<github Spreadsheet-XLSX-Reader-LibXML/issues
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

L<version>

L<Moose::Role>

L<Types::Standard>

=back

=head1 SEE ALSO

=over

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9