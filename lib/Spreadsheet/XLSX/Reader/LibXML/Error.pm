package Spreadsheet::XLSX::Reader::LibXML::Error;
use version; our $VERSION = qv('v0.18.2');

use Moose;
use Carp qw( cluck );
no warnings 'once';
@Spreadsheet::XLSX::Reader::LibXML::Error::CARP_NOT = qw(
	Capture::Tiny
);
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Str
		Bool
    );
use lib	'../../../../../lib',;
with 'Spreadsheet::XLSX::Reader::LibXML::LogSpace';
###LogSD	use Log::Shiras::TapWarn qw( re_route_warn restore_warn );
###LogSD	use Log::Shiras::Telephone;

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

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
			}else{
				###LogSD	my	$phone = Log::Shiras::Telephone->new(
				###LogSD			name_space => $self->get_log_space . '::error', );
				###LogSD		$phone->talk( level => 'debug', message => [
				###LogSD			"$_[1] line " .
				###LogSD			(((caller(2))[2])? ((caller(2))[2]) : ((caller(1))[2]) ) ] );
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

Spreadsheet::XLSX::Reader::LibXML::Error - Moose class for remembering the last error

=head1 SYNOPSIS
	
    #!/usr/bin/env perl
    $|=1;
    use MooseX::ShortCut::BuildInstance qw( build_instance );
    use Spreadsheet::XLSX::Reader::LibXML::Error;

    my  $action = build_instance(
            add_attributes =>{ 
                error_inst =>{
                    handles =>[ qw( error set_error clear_error set_warnings if_warn ) ],
                },
            },
			error_inst => Spreadsheet::XLSX::Reader::LibXML::Error->new(
                should_warn => 1,
                # should_warn => 0,# to turn off cluck when the error is set
            ),
        );
    print $action->dump;
          $action->set_error( "You did something wrong" );
    print $action->dump;
    print $action->error . "\n";
	
    ##############################################################################
    # SYNOPSIS Screen Output
    # 01: $VAR1 = bless( {
    # 02:             'error_inst' => bless( {
    # 03:                                 'should_warn' => 1,
    # 04:                                 'log_space' => 'Spreadsheet::XLSX::Reader::LogSpace'
    # 04:                             }, 'Spreadsheet::XLSX::Reader::Error' )
    # 05:         }, 'ANONYMOUS_SHIRAS_MOOSE_CLASS_1' );
    # 06: You did something wrong line 19
    # 07: at ../lib/Spreadsheet/XLSX/Reader/Error.pm line 28.
    # 08:    Spreadsheet::XLSX::Reader::Error::__ANON__('Spreadsheet::XLSX::Reader::Error=HASH(0x45e818)', 'You did something wrong') called at writer Spreadsheet::XLSX::Reader::Error::set_error of attribute error_string (defined at ../lib/Spreadsheet/XLSX/Reader/Error.pm line 28) line 9
    # 09:    Spreadsheet::XLSX::Reader::Error::set_error('Spreadsheet::XLSX::Reader::Error'=HASH(0x45e818)', 'You did something wrong') called at C:/strawberry/perl/site/lib/Moose/Meta/Method/Delegation.pm line 110
    # 10:    ANONYMOUS_SHIRAS_MOOSE_CLASS_1::set_error('ANONYMOUS_SHIRAS_MOOSE_CLASS_1=HASH(0x45e890)', 'You did something wrong') called at error_example.pl line 19
    # 11: $VAR1 = bless( {
    # 12:             'error_inst' => bless( {
    # 13:                                 'should_warn' => 1,
    # 14:                                 'error_string' => 'You did something wrong'
    # 15:                             }, 'Spreadsheet::XLSX::Reader::Error' )
    # 16:         }, 'ANONYMOUS_SHIRAS_MOOSE_CLASS_1' );
    # 17: You did something wrong
    ##############################################################################
    
=head1 DESCRIPTION

This L<Moose> class contains two L<attributes|Moose::Manual::Attributes>.  It is intended 
to be used through (by) L<delegation|Moose::Manual::Delegation> in other classes.  The first 
attribute is used to store the current error string.  The second, is set to turn on or off 
pushing the error string to STDERR when the first attribute is (re)set.

This class uses a role for full functional implementation.  Read the documentation 
for each of the role(s) as well as this documentation to gain a complete picture of this 
class.

=head2 with

These are attached role(s) for additional (re-used) functionality

=head3 L<Spreadsheet::XLSX::Reader::LibXML::LogSpace>

=head2 Attributes

Data passed to new when creating an instance.   For modification of 
these attributes see the listed 'attribute methods'. For more information on 
attributes see L<Moose::Manual::Attributes>.

=head3 error_string

=over

B<Definition:> This stores the most recent error string for recall later.  It 
also provides a hook to debug logging through 
L<Log::Shiras|https://github.com/jandrew/Log-Shiras>.

B<Default> undef (init_arg = undef)

B<Range> any string

B<attribute methods> Methods provided to adjust this attribute
		
=back

=head4 error

=over

B<Definition:> returns the currently stored error

=back

=head4 clear_error

=over

B<Definition:> clears the currently stored error

=back

=head4 set_error( $error_string )

=over

B<Definition:> sets the attribute to $error_string.

=back

=head3 should_warn

=over

B<Definition:> This determines if the package will L<cluck|CARP> and 
return a stack trace when the error_string attribute is set.

B<Default> 1 -> it will cluck

B<Range> Boolean values

B<attribute methods> Methods provided to adjust this attribute
		
=back

=head4 set_warnings( $bool )

=over

B<Definition:> Turn clucked warnings on or off

=back

=head4 if_warn

=over

B<Definition:> Returns the current setting of this attribute

=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> get clases in this package to return error numbers rather than 
error strings and then provide opportunity for this class to localize.

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

L<Moose>

L<Carp> - cluck

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<Types::Standard>

L<lib>

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

#########1#########2 main pod documentation end  5#########6#########7#########8#########9