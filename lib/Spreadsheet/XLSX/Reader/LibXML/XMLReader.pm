package Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use version; our $VERSION = qv('v0.32.2');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int				Str				HasMethods
		Bool
    );
use XML::LibXML::Reader;
use lib	'../../../../../lib',;
###LogSD	with 'Log::Shiras::LogSpace';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML::Types qw(
		IOFileType
	);

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has file_handle =>(
		isa			=> IOFileType,
		reader		=> 'get_file_handle',
		writer		=> 'set_file_handle',
		predicate	=> 'has_file_handle',
		clearer		=> 'clear_file_handle',
	);

has	error_inst =>(
		isa			=> 	HasMethods[qw(
							error set_error clear_error set_warnings if_warn
						) ],
		clearer		=> '_clear_error_inst',
		reader		=> '_get_error_inst',
		required	=> 1,
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
	);

has xml_reader =>(
	isa			=> 'XML::LibXML::Reader',
	reader		=> '_get_xml_parser',
	writer		=> '_set_xml_parser',
	predicate	=> '_has_xml_parser',
	clearer		=> '_clear_xml_parser',
	handles	=>{
		copy_current_node	=> 'copyCurrentNode',
		byte_consumed		=> 'byteConsumed',
		start_reading		=> 'read',
		next_element		=> 'nextElement',
		next_sibling		=> 'nextSibling',
		get_attribute		=> 'getAttribute',
		read_state			=> 'readState',
		node_name			=> 'name',
		node_value			=> 'value',
		has_value			=> 'hasValue',
		inner_xml			=> 'readInnerXml',
		node_depth			=> 'depth',
		is_empty			=> 'isEmptyElement',
		inner_xml			=> 'readInnerXml',
		has_attributes		=> 'hasAttributes',
		get_attribute_count	=> 'attributeCount',
		read_attribute		=> 'readAttributeValue',
		constant_value		=> 'ConstValue',
		move_to_first_att	=> 'moveToFirstAttribute',
		move_to_next_att	=> 'moveToNextAttribute',
		encoding			=> 'encoding',
		_go_to_the_end		=> 'finish',
		_close_the_sheet	=> 'close',
		next_sibling_element	=> 'nextSiblingElement',
	},
	trigger => \&_reader_init,
);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9


sub start_the_file_over{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space => $self->get_log_space . '::start_the_file_over', );
	###LogSD		$phone->talk( level => 'debug', message =>[ "Resetting the XML file" ] );
	$self->_go_to_the_end;
	$self->_close_the_sheet;
	#~ $self->_clear_xml_parser;
	$self->_clear_location;
	my $fh = $self->get_file_handle;
	$fh->seek( 0, 0 );
	$self->_set_xml_parser( XML::LibXML::Reader->new( IO => $fh ) );
	$self->start_reading;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _position_index =>(
		isa			=> Int,
		reader		=> 'where_am_i',
		writer		=> '_i_am_here',
		clearer		=> '_clear_location',
		predicate	=> 'has_position',
	);

has _read_unique_bits =>(
		isa		=> Bool,
		reader	=> '_get_unique_bits',
		writer	=> '_need_unique_bits',
		clearer	=> '_clear_read_unique_bits',
		default	=> 1,
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _reader_init{
	my( $self, $reader ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::_reader_init', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"loading any file specific settings", ] );
	
	# Set the file unique bits
	if( $self->_get_unique_bits ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Opening the file for the first time" ], );
		$self->_need_unique_bits( 0 );
		if( $self->can( '_load_unique_bits' ) ){
			###LogSD	$phone->talk( level => 'debug', message =>[ "Loading unique bits" ], );
			$self->_load_unique_bits;
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Finished loading unique bits" 			], );
		}
		$self->start_the_file_over;
	}else{
		###LogSD	$phone->talk( level => 'debug', message =>[ 
		###LogSD		"This is not the first time the file has been opened - don't seek unique" ], );
	}
	###LogSD	$phone->talk( level => 'debug', message => [ "Finished loading unique bits BLOCK" ], );
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::XMLReader::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"clearing the reader for log space:" . $self->get_log_space, ] );
	if( $self->_get_xml_parser ){
		#~ print "Disconnecting the sheet file handle from the parser\n";
		###LogSD	$phone->talk( level => 'debug', message =>[ "Disconnecting the file handle from the xml parser", ] );
		$self->_close_the_sheet;
		$self->_clear_xml_parser;
	}
	if( my $fh = $self->get_file_handle ){
		#~ print "Clearing file handle\n";
		###LogSD	$phone->talk( level => 'debug', message =>[ "Closing the system file handle", $self->dump(2) ] );
		$fh->close;
		###LogSD	$phone->talk( level => 'debug', message =>[ "Clearing the system file handle", $self->dump(2) ] );
		$self->clear_file_handle;
		###LogSD	$phone->talk( level => 'debug', message =>[ "Final self", $self->dump(2) ] );
	}
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader - LibXML::Reader base class for xlsx sheets

=head1 SYNOPSIS

	package MyPackage;
	use MooseX::StrictConstructor;
	use MooseX::HasDefaults::RO;
	extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
    
=head1 DESCRIPTION

B<This documentation is written to explain ways to extend this package.  To use the data 
extraction of Excel workbooks, worksheets, and cells please review the documentation for  
L<Spreadsheet::XLSX::Reader::LibXML>,
L<Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Spreadsheet::XLSX::Reader::LibXML::Cell>>

When setting worksheet file handles you can't reuse them again to re-open the same sheet 
since the last act of this package is to close the file handle before returning it.  You 
must also set both the file handle and the xml reader when building an instance of this 
class.

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