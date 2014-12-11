package Spreadsheet::XLSX::Reader::LibXML::Types;
use version; our $VERSION = qv('v0.26.2');
		
use strict;
use warnings;
use Type::Utils -all;
use Type::Library 0.046
	-base,
	-declare => qw(
		FileName					XMLFile						XLSXFile
		ParserType					Excel_number_0				EpochYear
		PassThroughType				CellID						PositiveNum
		NegativeNum					ZeroOrUndef					NotNegativeNum
						
	);
BEGIN{ extends "Types::Standard" };
#~ use Types::Standard  qw( Str InstanceOf Enum Num Any Maybe StrMatch );
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
#~ use lib	'../../../../lib',;
#~ ###LogSD	use Log::Shiras::Telephone;

#########1 Package Variables  3#########4#########5#########6#########7#########8#########9



#########1 Type Library       3#########4#########5#########6#########7#########8#########9

declare FileName,
	as Str,
    where{ -r $_ },
    message{ 
        ( $_ ) ? 
            "Could not find / read the file: $_" : 
            'No value passed to the file_name test' 
    };
	
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

my	$io_file_instance = InstanceOf[ 'IO::File' ];
declare XLSXFile,
	as Str|GlobRef|$io_file_instance,
	where{
		if( is_Str( $_ ) ){
			return $_ =~ /\.xlsx$/ and -r $_;
		}elsif( is_GlobRef( $_ ) ){
			return 1;
		}elsif( $io_file_instance->check( $_ ) ){
			return 1;
		}else{
			return 0;
		}
	},
	message{
		my $return;
		my $test = $_;
		if( !$test){
			$return = 'No value passed to the XMLFile test';
		}elsif( is_Str( $test ) ){
			$return = ( $test !~ /\.xml$/ ) ?
				"The string -$test- does not have an xml file extension" :
				"Could not find / read the file: $test" ;
		}else{
			$return = "|$test| is not a file handle I recognize";
		}
		return $return;
    };

declare ParserType, 
	as Enum[qw( reader )];#dom  sax

coerce ParserType,
	from Str,
	via{ lc( $_ ) };

declare EpochYear,
	as Int,
	where{ $_ == 1900 or $_ == 1904 };

declare PassThroughType,
	as Maybe[Any];

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


#########1 Excel Defined Converions     4#########5#########6#########7#########8#########9

declare_coercion Excel_number_0,
	to_type Any, from Maybe[Any],
	via{ $_ };

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9
	

#########1 Phinish            3#########4#########5#########6#########7#########8#########9
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::Types - A type library for the LibXML xlsx reader
    
=head1 DESCRIPTION

B<This documentation is written to explain ways to extend this package.  To use the data 
extraction of Excel workbooks, worksheets, and cells please review the documentation for  
L<Spreadsheet::XLSX::Reader::LibXML>,
L<Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Spreadsheet::XLSX::Reader::LibXML::Cell>>

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

#########1#########2 main pod documentation end  5#########6#########7#########8#########9