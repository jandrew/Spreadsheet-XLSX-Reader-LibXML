package Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings;
use version; our $VERSION = qv('v0.36.26');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int				HashRef			is_HashRef
    );
use Carp qw( confess );
use lib	'../../../../../../lib';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
with	'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData';

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub get_shared_string_position{
	my( $self, $position ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::get_shared_string_position', );
	if( !defined $position ){
		$self->set_error( "Requested shared string position required - none passed" );
		return undef;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Getting the sharedStrings position: $position" ] );
	
	#checking if the reqested position is too far
	if( $position > $self->_get_unique_count - 1 ){
		$self->set_error( "Asking for position -$position- (from 0) but the shared string " .
							"max cell position is: " . ($self->_get_unique_count - 1) );
		return undef;#  fail
	}
	
	# checking if the reqested position is stored
	if( $self->_has_last_position and $position == $self->_get_last_position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Already built the answer for position: $position", 
		###LogSD		$self->_get_last_position_ref						] );
		return $self->_get_last_position_ref;
	}
	
	# Initiate the read and reset the file if needed
	if( $self->has_position and $self->where_am_i > $position ){
		$self->start_the_file_over;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Finished resetting the file" ] );
	}
	if( !$self->has_position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the first cell" ] );
		my $found_it;
		eval '$found_it = $self->next_element( "si" )';
		if( $@ ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Found an unexpected end of file: ", $@] );
			$self->set_error( 'libxml2 error message' . $@ );
			$self->_set_unique_count( 0 );
			$self->_i_am_here( 0 );
			return undef;
		}elsif( defined $found_it and $found_it < 1 ){
			$self->set_error( "No strings stored in the sharedStrings file" );
			return undef;
		}
		$self->_i_am_here( 0 );
	}
	
	# Advance to the proper position
	while( $self->where_am_i < $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the next cell: " . ($self->where_am_i + 1) ] );
		my $found_it;
		eval '$found_it = $self->next_element( "si" )';
		if( $@ ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Found an unexpected end of file: ", $@] );
			$self->set_error( 'libxml2 error message' . $@ );
			$self->_set_unique_count( $self->where_am_i + 1 );
			$self->_i_am_here( 0 );
			return undef;
		}elsif( defined $found_it and $found_it < 1 ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Found an unexpected end of file: $found_it" ] );
			$self->set_error( "Unexpected end of file found" );
			$self->_set_unique_count( $self->where_am_i + 1 );
			$self->_i_am_here( 0 );
			return undef;
		}
		$self->_i_am_here( $self->where_am_i + 1 );
	}
	
	# Read the data
	$self->_set_last_position( $position );
	my $init_ref = $self->parse_element;
	$self->_i_am_here( $position + 1 );
	###LogSD	$phone->talk( level => 'trace',  message =>[
	###LogSD		"Element parse resulted in:", $init_ref ] );
	if( is_HashRef( $init_ref ) ){
		###LogSD	$phone->talk( level => 'trace',  message =>[
		###LogSD		"This is a hash ref" ] );
		if( exists $init_ref->{list} ){
			my ( $raw_text, $rich_text );
			for my $element( @{$init_ref->{list}} ){
				push( @$rich_text, length( $raw_text ), $element->{rPr} ) if exists $element->{rPr};
				$raw_text .= $element->{t}->{raw_text};
			}
			@$init_ref{qw( raw_text rich_text )} = ( $raw_text, $rich_text  );
			delete $init_ref->{list};
		}else{
			$init_ref = $init_ref->{t};
		}
	}else{
		$self->set_error( "Unable to parse the shared string position: $position" );
		return undef;
	}
	$self->_set_last_position_ref( $init_ref );
	###LogSD	$phone->talk( level => 'trace',  message =>[
	###LogSD		"Final element rearranging result:", $init_ref ] );
	return $init_ref;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _unique_count =>(
	isa			=> Int,
	writer		=> '_set_unique_count',
	reader		=> '_get_unique_count',
	clearer		=> '_clear_unique_count',
);

has _last_position =>(
		isa		=> Int,
		writer	=> '_set_last_position',
		reader	=> '_get_last_position',
		predicate => '_has_last_position',
		trigger	=> \&_clear_last_position_ref,
	);

has _last_position_ref =>(
		isa		=> HashRef,
		writer	=> '_set_last_position_ref',
		reader	=> '_get_last_position_ref',
		clearer => '_clear_last_position_ref',
		predicate => '_has_last_position_ref',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _load_unique_bits{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_load_unique_bits', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the sharedStrings unique bits" ] );
	
	if( $self->node_name eq 'sst' or $self->next_element('sst') ){
		my $sst_list= $self->parse_element( 0 );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"parsed sst list to:", $sst_list ] );
		my $unique_count = $sst_list->{uniqueCount} // 0;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Loading unique count: $unique_count" ] );
		$self->_set_unique_count( $unique_count );
		return undef;
	}else{
		$self->set_error( "No 'sst' element found - can't parse this as a shared strings file" );
		$self->_clear_unique_count;
		$self->_clear_count;
		return undef;
	}
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::XMLReader::SharedStrings - A LibXML::Reader sharedStrings base class

=head1 SYNOPSIS

	#!/usr/bin/env perl
	use Data::Dumper;
	use MooseX::ShortCut::BuildInstance qw( build_instance );
	use Spreadsheet::XLSX::Reader::LibXML::Error;
	use Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings;

	my $file_instance = build_instance(
	    package      => 'SharedStringsInstance',
	    superclasses => ['Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings'],
	    file         => 'sharedStrings.xml',
	    error_inst   => Spreadsheet::XLSX::Reader::LibXML::Error->new,
	);
	print Dumper( $file_instance->get_shared_string_position( 3 ) );
	print Dumper( $file_instance->get_shared_string_position( 12 ) );

	#######################################
	# SYNOPSIS Screen Output
	# 01: $VAR1 = {
	# 02:     'raw_text' => ' '
	# 03: };
	# 04: $VAR1 = {
	# 05:     'raw_text' => 'Superbowl Audibles'
	# 06: };
	#######################################
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module when writing your 
own excel parser or extending this package.  To use the general package for excel 
parsing out of the box please review the documentation for L<Workbooks
|Spreadsheet::XLSX::Reader::LibXML>, L<Worksheets
|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>.

This class is written to extend L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>.  
It addes to that functionality specifically to read the sharedStrings.xml sub file 
zipped into an .xlsx file.  It does not provide connection to other file types or 
even the elements from other files that are related to this file.  This POD only 
describes the functionality incrementally provided by this module.  For an overview 
of sharedStrings.xml reading see L<Spreadsheet::XLSX::Reader::LibXML::SharedStrings>

=head2 Methods

These are the methods provided by this module.

=head3 get_shared_string_position( $positive_int )

=over

B<Definition:> This returns the xml converted to a deep perl data structure from the 
defined 'si' position.

B<Accepts:> $positive_int ( a positive integer )

B<Returns:> a deep perl data structure built from the xml at 'si' position $positive_int

=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Nothing yet

=back

=head1 AUTHOR

=over

Jed Lund

jandrew@cpan.org

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

#########1#########2 main pod documentation end   5#########6#########7#########8#########9