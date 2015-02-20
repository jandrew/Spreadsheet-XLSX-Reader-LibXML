package Spreadsheet::XLSX::Reader::LibXML::Types;
use version; our $VERSION = qv('v0.34.4');
		
use strict;
use warnings;
use Type::Utils -all;
use Type::Library 1.000
	-base,
	-declare => qw(
		XMLFile						XLSXFile					ParserType										
		NegativeNum					ZeroOrUndef					NotNegativeNum
		IOFileType					ErrorString					SubString
		CellID						PositiveNum					Excel_number_0
	);
use IO::File;
BEGIN{ extends "Types::Standard" };
my $try_xs =
		exists($ENV{PERL_TYPE_TINY_XS}) ? !!$ENV{PERL_TYPE_TINY_XS} :
		exists($ENV{PERL_ONLY})         ?  !$ENV{PERL_ONLY} :
		1;
if( $try_xs and exists $INC{'Type/Tiny/XS.pm'} ){
	eval "use Type::Tiny::XS 0.010";
	if( $@ ){
		die "You have loaded Type::Tiny::XS but versions prior to 0.010 will cause this module to fail";
	}
}

#########1 Package Variables  3#########4#########5#########6#########7#########8#########9



#########1 Type Library       3#########4#########5#########6#########7#########8#########9
	
declare XMLFile,
	as Str,
	where{ $_ =~ /\.xml$/ and -r $_},
	message{
		( $_ !~ /\.xml$/ ) ?
			"The string -$_- does not have an xml file extension" :
		( !-r $_ ) ?
			"Could not find / read the file: $_" :
			'No value passed to the XMLFile test';
    };

declare XLSXFile,
	as Str,
	where{ $_ =~ /\.xlsx$/ and -r $_ },
	message{
		my $test = $_;
		my $return =
			( !defined $test ) ?
				"Empty filename" :
			( ref $test ) ?
				"'" . $test . "' is not a string value" :
			( $test !~ /\.xlsx$/ ) ?
				"The string -$test- does not have an xlsx file extension" :
			( -r $test) ?
				"Could not find / read the file: $test" :
				"Unmanageable value '" . ($test ? $test : '' ) . "' passed" ;
		return $return;
    };
	
declare IOFileType,
	as InstanceOf[ 'IO::File' ];
	
coerce IOFileType,
	from GlobRef,
	via{  bless $_, 'IO::File' };
	
coerce IOFileType,
	from XLSXFile,
	via{  IO::File->new( $_, 'r' ); };
	
coerce IOFileType,
	from XMLFile,
	via{  IO::File->new( $_, 'r' ); };

declare ParserType,
	as Enum[qw( reader )];#dom  sax

coerce ParserType,
	from Str,
	via{ lc( $_ ) };

declare CellID,
	as StrMatch[ qr/^[A-Z]{1,3}[1-9]\d*$/ ];
	
declare PositiveNum,
	as Num,
	where{ $_ > 0 };

declare NegativeNum,
	as Num,
	where{ $_ < 0 };
	
declare ZeroOrUndef,
	as Maybe[Num],
	where{ !$_ };
	
declare NotNegativeNum,
	as Num,
	where{ $_ > -1 };

declare SubString,
	as Str;

declare ErrorString,
	as SubString,
	where{ $_ !~ /\)\n;/ };# Since the coercion removes them
	
coerce SubString,
	from Object,
	via{ 
	my	$object = $_;
		if( $object->can( 'as_string' ) ){
			return $object->as_string;
		}elsif( $object->can( 'message' ) ){
			return $object->message;
		}
		return $object;
	};
	
coerce ErrorString,
	from SubString->coercibles,
	via{
	my	$tmp = to_SubString($_);
		$tmp =~ s/\)\n;/\);/g;
		return $tmp;
	};


#########1 Excel Defined Converions     4#########5#########6#########7#########8#########9

declare_coercion Excel_number_0,
	to_type Any, from Maybe[Any],
	via{ $_ };

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9
	

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

__PACKAGE__->meta->make_immutable;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::Types - A type library for the LibXML xlsx reader
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module.  To use the general 
package for excel parsing out of the box please review the documentation for L<Workbooks
|Spreadsheet::XLSX::Reader::LibXML>, L<Worksheets
|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>.

This is a L<Type::Library|Type::Tiny::Manual::Libraries> for this package.

=head1 TYPES

=head2 XMLFile

This type checks that the value is a readable file (full path - no file find magic 
used) with an \.xml extention

=head3 coercions

none

=head2 XLSXFile

This type checks that the value is a readable file (full path - no file find magic 
used)  with an \.xlsx extention

=head3 coercions

none

=head2 ParserType

For now this type checks that the parser type string == 'reader'.  As future parser 
types are added to the package I will update this type.

=head3 coercions

=over

B<Str:> this will lower case any other version of reader (Reader| READER) to get it to pass

=back

=head2 NegativeNum

ZeroOrUndef					NotNegativeNum
		IOFileType					ErrorString					SubString
		CellID						PositiveNum					Excel_number_0
=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> The ErrorString type tests still needs a 'fail' case

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

#########1#########2 main pod documentation end  5#########6#########7#########8#########9
