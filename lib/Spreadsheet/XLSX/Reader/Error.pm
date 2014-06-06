package Spreadsheet::XLSX::Reader::Error;
use version; our $VERSION = version->declare("v0.1_1");

use Moose;
use Carp qw( cluck );
no warnings 'once';
@Spreadsheet::XLSX::Reader::Error::CARP_NOT = qw(
	Capture::Tiny
);
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Str
		Bool
    );
use lib	'../../../../lib',;
###LogSD	use Log::Shiras::TapWarn qw( re_route_warn restore_warn );

has error_string =>(
		isa		=> Str,
		clearer	=> 'clear_error',
		reader	=> 'error',
		writer	=> 'set_error',
		trigger	=> sub{
			my ( $self ) = @_;
			if( $self->if_warn ){
				###LogSD	re_route_warn();
				cluck "$_[1] line " .
					(((caller(2))[2])? ((caller(2))[2]) : ((caller(1))[2]) ). "\n";
				###LogSD	restore_warn;
			}
		},
		init_arg => undef,
	);

has should_warn =>(
		isa		=> Bool,
		default	=> 1,
		writer	=> 'set_warnings',
		reader	=> 'if_warn',
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9

__END__

=head1 NAME

Spreadsheet::XLSX::Reader::Error - Moose class for remembering the last error
    
=head1 DESCRIPTION

This role contains two attributes.  The first is used to store the current error string.  
The second, is set to turn on or off pushing the error string to STDERR.

=head1 SYNOPSIS
	
	#!perl
	$|=1;
	use MooseX::ShortCut::BuildInstance qw( build_instance );
	use Spreadsheet::XLSX::Reader::Error;

	my 	$action = build_instance(
			add_attributes =>{ 
				error_inst =>{
					handles =>[ qw( error set_error clear_error set_warnings if_warn ) ],
				},
			},
			error_inst => Spreadsheet::XLSX::Reader::Error->new(
				should_warn => 1,
				# should_warn => 0,# to turn off cluck when the error is set
			),
		);
	print	$action->dump;
			$action->set_error( "You did something wrong" );
	print	$action->dump;
	print	$action->error . "\n";
	
	##############################################################################
	# SYNOPSIS Screen Output
	# 01: $VAR1 = bless( {
	# 02:                  'error_inst' => bless( {
	# 03:                                           'should_warn' => 1
	# 04:                                         }, 'ANONYMOUS_SHIRAS_MOOSE_CLASS_1' )
	# 05:                }, 'ANONYMOUS_SHIRAS_MOOSE_CLASS_2' );
	# 06: You did something wrong line 19
 	# 07: at ../lib/Spreadsheet/XLSX/Reader/Error.pm line 22.
	# 08:        Spreadsheet::XLSX::Reader::Error::__ANON__('ANONYMOUS_SHIRAS_MOOSE_CLASS_1=HASH(0x45e818)', 'You did something wrong') called at writer Spreadsheet::XLSX::Reader::Error::set_error of attribute error_string (defined at ../lib/Spreadsheet/XLSX/Reader/Error.pm line 28) line 9
 	# 09:        Spreadsheet::XLSX::Reader::Error::set_error('ANONYMOUS_SHIRAS_MOOSE_CLASS_1=HASH(0x45e818)', 'You did something wrong') called at C:/strawberry/perl/site/lib/Moose/Meta/Method/Delegation.pm line 110
 	# 10:        ANONYMOUS_SHIRAS_MOOSE_CLASS_3::set_error('ANONYMOUS_SHIRAS_MOOSE_CLASS_2=HASH(0x45e890)', 'You did something wrong') called at error_example.pl line 19
 	# 11: $VAR1 = bless( {
  	# 12:                 'error_inst' => bless( {
 	# 13:                                           'should_warn' => 1,
 	# 14:                                           'error_string' => 'You did something wrong'
 	# 15:                                         }, 'ANONYMOUS_SHIRAS_MOOSE_CLASS_1' )
 	# 16:                }, 'ANONYMOUS_SHIRAS_MOOSE_CLASS_2' );
 	# 17: You did something wrong
	##############################################################################

=head2 Attributes

Data passed to new when creating an instance (parser).  For modification of 
these attributes see the listed L</Methods> of the instance.

=head3 error_string

=over

B<Definition:> This stores the most recent error string for recall later.  It 
also provides a link to debug logging as needed.

B<Default> undef (init_arg = undef)

B<Range> any string
		
=back

=head3 should_warn

=over

B<Definition:> This determins if the package will L<cluck|CARP> and 
return a stack trace when the error_string attribute is set.

B<Default> 1 -> it will cluck

B<Range> Bool
		
=back

=head2 Methods

These include methods to adjust attributes.
		
=head3 set_warnings( $bool )

=over

B<Definition:> Turn clucked warnings on or off

B<Accepts:> Boolean values

B<Returns:> nothing

=back
		
=head3 if_warn

=over

B<Definition:> Check the current should_warn attribute value

B<Accepts:> Nothing

B<Returns:> $bool representing the current should_warn attribute value

=back
		
=head3 error

=over

B<Definition:> Returns the currently stored error string

B<Accepts:> Nothing

B<Returns:> $error representing the current error_string attribute value

=back

=head3 _set_error( $error_string )

=over

B<Definition:> internal method to set (update) the current error string

B<Accepts:> $error_string

B<Returns:> Nothing

=back

=head3 _clear_error

=over

B<Definition:> internal method to clear the current error string

B<Accepts:> Nothing

B<Returns:> Nothing (string is cleared)

=back

=head1 SUPPORT

=over

L<github Spreadsheet-XLSX-Reader/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader/issues>

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

L<Carp> - cluck

L<Types::Standard>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::XLSX::Reader::TempFilter>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9