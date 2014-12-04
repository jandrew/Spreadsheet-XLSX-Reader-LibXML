package Spreadsheet::XLSX::Reader::LibXML;
use version 0.77; our $VERSION = qv('v0.22.2');

use 5.010;
use	List::Util 1.33;
use	Moose;
use	MooseX::StrictConstructor;
use	MooseX::HasDefaults::RO;
use	Carp qw( confess );
use	Archive::Zip;
use	OLE::Storage_Lite;
use	File::Temp;
#~ $File::Temp::DEBUG = 1;
#~ use	Data::Dumper;
use Types::Standard qw(
 		InstanceOf			Str       		StrMatch
		Enum				HashRef			ArrayRef
		CodeRef				Int				HasMethods
		Bool
    );
use	MooseX::ShortCut::BuildInstance 1.028 qw( build_instance should_re_use_classes );
should_re_use_classes( 1 );
use lib	'../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
with 	'Spreadsheet::XLSX::Reader::LibXML::LogSpace';
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles;
use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use	Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet;
use	Spreadsheet::XLSX::Reader::LibXML::Types qw( XLSXFile ParserType );

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9

my	$parser_modules ={
		reader =>{
			sharedStrings =>{
				superclasses	=> ['Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings'],
				attributes		=> [qw( error_inst )],
				store			=> '_set_shared_strings_instance',
				package			=> 'SharedStringsInstance',
			},
			styles =>{
				superclasses			=> ['Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles'],
				attributes				=> [qw( epoch_year error_inst )],
				add_roles_in_sequence	=> [qw( default_format_list format_string_parser )],
				store					=> '_set_styles_instance',
				package					=> 'StylesInstance',
			},
			worksheet =>{
				superclasses	=> ['Spreadsheet::XLSX::Reader::LibXML::XMLReader::Worksheet'],
				store			=> '_set_worksheet_superclass',
			},
		},
	};

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has	error_inst =>(
		isa			=> 	HasMethods[qw(
							error set_error clear_error set_warnings if_warn
						) ],
		clearer		=> '_clear_error_inst',
		reader		=> 'get_error_inst',
		required	=> 1,
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
		default => sub{ Spreadsheet::XLSX::Reader::LibXML::Error->new( should_warn => 0 ) },
	);

has file_name =>(
		isa			=> XLSXFile,
		writer		=> 'set_file_name',
		clearer		=> '_clear_file_name',
		predicate	=> 'has_file_name',
		trigger		=> \&_build_file,
	);
	
has file_creator =>(
		isa		=> Str,
		reader	=> 'creator',
		writer	=> '_set_creator',
		clearer	=> '_clear_creator',
	);
	
has file_modified_by =>(
		isa		=> Str,
		reader	=> 'modified_by',
		writer	=> '_set_modified_by',
		clearer	=> '_clear_modified_by',
	);
	
has file_date_created =>(
		isa		=> StrMatch[qr/^\d{4}\-\d{2}\-\d{2}/],
		reader	=> 'date_created',
		writer	=> '_set_date_created',
		clearer	=> '_clear_date_created',
	);
	
has file_date_modified =>(
		isa		=> StrMatch[qr/^\d{4}\-\d{2}\-\d{2}/],
		reader	=> 'date_modified',
		writer	=> '_set_date_modified',
		clearer	=> '_clear_date_modified',
	);

has sheet_parser =>(
		isa		=> ParserType,
		writer	=> 'set_parser_type',
		reader	=> 'get_parser_type',
		default	=> 'reader',
		coerce	=> 1,
	);

has count_from_zero =>(
		isa		=> Bool,
		reader	=> 'counting_from_zero',
		writer	=> 'set_count_from_zero',
		default	=> 1,
	);
	
has file_boundary_flags =>(
		isa			=> Bool,
		reader		=> 'boundary_flag_setting',
		writer		=> 'change_boundary_flag',
		default		=> 1,
		required	=> 1,
	);

has empty_is_end =>(
		isa		=> Bool,
		writer	=> 'set_empty_is_end',
		reader	=> 'is_empty_the_end',
		default	=> 0,
	);

has from_the_edge =>(
		isa		=> Bool,
		reader	=> '_starts_at_the_edge',
		writer	=> 'set_from_the_edge',
		default	=> 1,
	);

has default_format_list =>(
		isa		=> Str,
		writer	=> 'set_default_format_list',
		reader	=> 'get_default_format_list',
		default	=> 'Spreadsheet::XLSX::Reader::LibXML::FmtDefault',
	);

has format_string_parser =>(
		isa		=> Str,
		writer	=> 'set_format_string_parser',
		reader	=> 'get_format_string_parser',
		default	=> 'Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings',
	);

has group_return_type =>(
		isa		=> Enum[qw( unformatted value instance )],
		reader	=> 'get_group_return_type',
		writer	=> 'set_group_return_type',
		default	=> 'instance',
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub parse{

    my ( $self, $file_name, $formatter ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::parse', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at parse for: $file_name",
	###LogSD			(($formatter) ? "with formatter: $formatter" : '') ] );
	$self->set_format_string_parser( $formatter ) if $formatter;
	$self->set_file_name( $file_name );
	return ( $self->has_file_name ) ? $self : undef;
}

sub worksheets{

    my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::worksheets', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			'Attempting to build all worksheets: ', $self->get_worksheet_names ] );
	my	@worksheet_array;
	while( my $worksheet_object = $self->worksheet ){
	#~ for my $worksheet_name ( @worksheet_array ){
		#~ my	$worksheet_object = $self->worksheet( $worksheet_name );
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		'Built worksheet: ' .  $worksheet_object->get_name ] );
		push @worksheet_array, $worksheet_object;#$self->worksheet( $worksheet_name );
	}
	###LogSD	$phone->talk( level => 'trace', message =>[
	###LogSD		'sending worksheet array: ',@worksheet_array ] );
	return @worksheet_array;
}

sub worksheet{

    my ( $self, $worksheet_name ) = @_;
	my ( $next_position );
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::worksheet', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			"Arrived at (build a) worksheet with: ", $worksheet_name ] );
	confess "No file loaded yet" if !$self->has_file_name;
	# Handle an implied 'next sheet'
	if( !$worksheet_name ){
		my $worksheet_position = $self->_get_current_worksheet_position;
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"No worksheet name passed", 
		###LogSD		((defined $worksheet_position) ? "Starting after position: $worksheet_position" : '')] );
		$next_position = ( !$self->in_the_list ) ? 0 : ($self->_get_current_worksheet_position + 1);
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"No worksheet name passed", "Attempting position: $next_position" ] );
		if( $next_position >= $self->number_of_sheets ){
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD		"Reached the end of the worksheet list" ] );
			return undef;
		}
		$worksheet_name = $self->worksheet_name( $next_position );
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Now attempting to build the worksheet named: $worksheet_name", ] );
	}
	
	# build the worksheet
	my	$worksheet_info = $self->_get_worksheet_info( $worksheet_name );
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		'Returned worksheet info:', $worksheet_info, ] );
	confess "No worksheet info for: $worksheet_name" if !exists $worksheet_info->{sheet_position};
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"Building the next worksheet with:",
	###LogSD		{
	###LogSD			superclasses	=> $self->_get_worksheet_superclass,
	###LogSD			log_space 		=> $self->get_log_space . "::Worksheet",
	###LogSD			sheet_name		=> $worksheet_name,
	###LogSD			workbook_instance => '(self)',
	###LogSD			%$worksheet_info,
	###LogSD		}										] );
	my	$worksheet = 	build_instance(
							superclasses		=> $self->_get_worksheet_superclass,
							package				=> 'WorksheetInstance',
							log_space 			=> $self->get_log_space . "::Worksheet",
							sheet_name			=> $worksheet_name,
							workbook_instance	=> $self,
							error_inst			=> $self->get_error_inst,
							%{$worksheet_info}, 
						);
	if( $worksheet ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Successfully loaded: $worksheet_name",] );
		$self->_set_current_worksheet_position( $worksheet->position );
		return $worksheet;
	}else{
		$self->set_error( "Failed to build the object for worksheet: $worksheet_name" );
		return undef;
	}
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _epoch_year =>(
		isa		=> Enum[qw( 1900 1904 )],
		writer	=> '_set_epoch_year',
		reader	=> 'get_epoch_year',
		default	=> 1900,
	);
	
has _shared_strings_instance =>(
		isa			=> HasMethods[ 'get_shared_string_position' ],
		predicate	=> '_has_shared_strings_file',
		writer		=> '_set_shared_strings_instance',
		reader		=> '_get_shared_strings_instance',
		clearer		=> '_clear_shared_strings',
		handles		=>{
			'get_shared_string_position' => 'get_shared_string_position',
			_demolish_shared_strings => 'DEMOLISH',
		},
	);
	
has _styles_instance =>(
		isa			=> HasMethods[qw( get_format_position )],
		writer		=> '_set_styles_instance',
		reader		=> '_get_styles_instance',
		clearer		=> '_clear_styles',
		predicate	=> '_has_styles_file',
		handles		=>{
			get_format_position	=> 'get_format_position',
			set_defined_excel_format_list => 'set_defined_excel_format_list',
			change_output_encoding => 'change_output_encoding',
			get_date_behavior => 'get_date_behavior',
			set_date_behavior => 'set_date_behavior',
			parse_excel_format_string => 'parse_excel_format_string',
			_demolish_styles => 'DEMOLISH',
		},
	);

has _calc_chain_instance =>(
	isa	=> 	HasMethods[qw( get_calc_chain_position )],
	writer	=>'_set_calc_chain_instance',
	reader	=>'_get_calc_chain_instance',
	clearer	=> '_clear_calc_chain',
	predicate => '_has_calc_chain_file',
	handles =>{
		_demolish_calc_chain => 'DEMOLISH',
	},
);

has '_temp_dir' =>(
	isa		=> 'File::Temp::Dir',
	writer	=> '_set_temp_dir',
	reader	=> '_get_temp_dir_object',
	clearer	=> '_clear_temp_dir',
	handles	=>{
		_get_temp_dir => 'dirname',
	},
);

has _worksheet_list =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		writer	=> '_set_worksheet_list',
		clearer	=> '_clear_worksheet_list',
		reader	=> 'get_worksheet_names',
		handles	=>{
			worksheet_name => 'get',
		},
		default	=> sub{ [] },
	);

has _worksheet_lookup =>(
		isa		=> HashRef,
		traits	=> ['Hash'],
		writer	=> '_set_worksheet_lookup',
		clearer	=> '_clear_worksheet_lookup',
		reader	=> '_get_worksheet_lookup',
		handles	=>{
			_get_worksheet_info => 'get',
			number_of_sheets	=> 'count',
		},
		default	=> sub{ {} },
	);

has _current_worksheet_position =>(
		isa			=> Int,
		writer		=> '_set_current_worksheet_position',
		reader		=> '_get_current_worksheet_position',
		clearer		=> 'start_at_the_beginning',
		predicate	=> 'in_the_list',
	);
	
has _worksheet_superclass =>(
		isa		=> ArrayRef,
		clearer	=> '_clear_worksheet_superclass',
		writer	=> '_set_worksheet_superclass',
		reader	=> '_get_worksheet_superclass',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _build_file{

    my ( $self, $file_name ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_build_file', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			'Arrived at _build_file for: ', $file_name ] );
	$self->_clear_shared_strings;
	$self->_clear_calc_chain;
	$self->_clear_styles;
	$self->_clear_temp_dir;
	$self->_clear_worksheet_list;
	$self->_clear_worksheet_lookup;
	$self->_clear_creator;
	$self->_clear_modified_by;
	$self->_clear_date_created;
	$self->_clear_date_modified;
	$self->clear_error;
	$self->start_at_the_beginning;
	my	$message;
    if ( !-e $file_name ){
		$message = "Can't locate -$file_name-";
	}elsif( -z $file_name ) {
		$message = "There is nothing in the file -$file_name-";
	}
	if( $message ){
		$self->_set_error( $message );
		$self->_clear_file_name;
		return;
	}

    # Check for xls or encrypted OLE files.
    my $ole_file = $self->_check_if_ole_file( $file_name );
    return( undef ) if $ole_file;

    # Create a, locally scoped, temp dir to unzip the XLSX file into.
	my	$temp_dir = File::Temp->newdir( DIR => undef );
    $self->_set_temp_dir( $temp_dir );
	
    # Create an Archive::Zip object to unzip the XLSX file.
    my $zip_file = Archive::Zip->new();

    # Read the XLSX zip file and catch any errors.
    eval { $zip_file->read( $file_name ) };
    if ( $@ ) {
		$self->_set_error( "File has zip error(s): " . join( ' ~|~ ', $@ ) );
		return undef;
	}

    # Extract the XML files from the XLSX zip.
    $zip_file->extractTree( '', $self->_get_temp_dir . '/' );
	###LogSD	$phone->talk( level => 'trace', message =>[ "Zip file: ", $zip_file ] );
	my	$intermediate = $self->_get_temp_dir;
	###LogSD	$phone->talk( level	=> 'trace', message =>[
	###LogSD		"Temp dir: $intermediate", "Temp Dir contains: ", <$intermediate/*> ] );
	
	# Load general workbook information to this instance
	my	$workbook_file = $self->_get_temp_dir . '/xl/workbook.xml';
	###LogSD	$phone->talk( level => 'debug', message =>[ "Loading workbook file: $workbook_file"	] );
	my ( $rel_lookup, $id_lookup ) = $self->_load_workbook_file( $workbook_file );
	return undef if !$rel_lookup;
	
	# Load the workbook rels file
	$workbook_file = $self->_get_temp_dir . '/xl/_rels/workbook.xml.rels';
	###LogSD	$phone->talk( level => 'debug', message =>[ "Loading _rels file: $workbook_file"	] );
	my ( $pivot_lookup ) = $self->_load_rels_workbook_file( $rel_lookup, $workbook_file );
	return undef if !$pivot_lookup;
	
	# Load the docProps file
	$workbook_file = $self->_get_temp_dir . '/docProps/core.xml';
	###LogSD	$phone->talk( level => 'debug', message =>[ "Loading _doc_props file: $workbook_file"	] );
	$self->_load_doc_props_file( $workbook_file );
	#~ my $wait = <>;
	# Build the instances for all the shared files (data for sheets shared across worksheets)
	if( exists $parser_modules->{ $self->get_parser_type } ){
		my	$result = 	$self->_set_shared_worksheet_files(
							$parser_modules->{ $self->get_parser_type }
						);
		return undef if !$result;
	}else{
		$self->_set_error( "No definitions for the sheet parser type: " .
								$self->get_parser_type );
		return undef;
	}
	return $self;
}

sub _check_if_ole_file {

    my ( $self, $file_name ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_check_if_ole_file', );
	###LogSD		$phone->talk( level => 'info', message =>[
	###LogSD			'Arrived at _check_if_ole_file for: ', $file_name, ], );
    my	$ole = OLE::Storage_Lite->new( $file_name );
    my	$pps = $ole->getPpsTree();

    # If getPpsTree() failed then this isn't an OLE file.
    return if !$pps;

    # Loop through the PPS children below the root.
	my	$message;
    for my $child_pps ( @{ $pps->{Child} } ) {

        my 	$pps_name = OLE::Storage_Lite::Ucs2Asc( $child_pps->{Name} );
		
        # Match an Excel xls file.
        if ( $pps_name eq 'Workbook' || $pps_name eq 'Book' ) {
			$message = "File is xls not an xlsx file: $file_name";
			last;
        }elsif( $pps_name eq 'EncryptedPackage' ) {
			$message = "File is encrypted an encrypted xlsx file: $file_name";
			last;
        }
    }
	
	#Handle result
	if( $message ){
		$self->_set_error( $message );
	}else{
		###LogSD	$phone->talk( level => 'warn', message =>[
		###LogSD		'The OLE test passed (negative) for: ', $file_name, ], );
		return undef;
	}
	return 1;
}

sub _load_workbook_file{
	my( $self, $new_file ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_load_workbook_file', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building the DOM for the workbook file: $new_file" ] );
	my	$dom = XML::LibXML->load_xml( location => $new_file );
	my ( $list, $sheet_ref, $rel_lookup, $id_lookup );
	my	$position = 0;
	my ( $setting_node ) = $dom->getElementsByTagName( 'workbookPr' );
	$self->_set_epoch_year( 1904 ) if $setting_node->getAttribute( 'date1904' );
	for my $sheet ( $dom->getElementsByTagName( 'sheet' ) ){
		my	$sheet_name = $sheet->getAttribute( 'name' );
		push @$list, $sheet_name;
		@{$sheet_ref->{$sheet_name}}{ 'sheet_id', 'sheet_rel_id', 'sheet_position' } = (
				$sheet->getAttribute( 'sheetId' ),
				$sheet->getAttribute( 'r:id' ),
				$position++,
		);
		$rel_lookup->{$sheet->getAttribute( 'r:id' )} = $sheet_name;
		$id_lookup->{$sheet->getAttribute( 'sheetId' )} = $sheet_name;
	}
	for my $sheet ( $dom->getElementsByTagName( 'pivotCache' ) ){
		my	$sheet_id = $sheet->getAttribute( 'cacheId' );
		my	$rel_id = $sheet->getAttribute( 'r:id' );
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Sheet ID: $sheet_id", "Rel ID: $rel_id", ] );
		$rel_lookup->{$rel_id} = $sheet_id;
		$id_lookup->{$sheet_id} = $rel_id;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Sheet list: ", $list,
	###LogSD		"Worksheet lookup:", $sheet_ref,
	###LogSD		"rel lookup:", $rel_lookup,
	###LogSD		"id lookup:", $id_lookup,		] );
	if( !$list ){
		$self->_set_error( "No worksheets identified in this workbook" );
		return undef;
	}
	$self->_set_worksheet_list( $list );
	$self->_set_worksheet_lookup( $sheet_ref );
	return( $rel_lookup, $id_lookup );
}

sub _load_rels_workbook_file{
	my( $self, $rel_lookup, $new_file ) = @_;
	my ( $pivot_lookup, );
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_load_rels_workbook_file', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building the DOM for the rels/workbook file: $new_file",
	###LogSD			"With rel_lookup:", $rel_lookup								] );
	my	$dom = XML::LibXML->load_xml( location => $new_file );
	my	$sheet_ref = $self->_get_worksheet_lookup;
	my	$found_file_names = 0;
	for my $sheet ( $dom->getElementsByTagName( 'Relationship' ) ){
		my	$rel_ID = $sheet->getAttribute( 'Id' );
		if( exists $rel_lookup->{$rel_ID} ){
			my	$file_name = $self->_get_temp_dir . '/xl/' . $sheet->getAttribute( 'Target' );
				$file_name =~ s/\\/\//g;
			if( $file_name =~ /worksheets/ ){
				$sheet_ref->{$rel_lookup->{$rel_ID}}->{file_name} = $file_name;
				$found_file_names = 1;
			}else{
				$pivot_lookup->{$rel_ID} = $file_name;
			}
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Worksheet lookup:", $sheet_ref,
	###LogSD		"Pivot lookup:", $pivot_lookup	] );
	if( !$found_file_names ){
		$self->_set_error( "Couldn't find any file names for the sheets" );
		return undef;
	}
	$self->_set_worksheet_lookup( $sheet_ref );
	return $pivot_lookup;
}

sub _load_doc_props_file{
	my( $self, $new_file ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_load_doc_props_file', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building the DOM for the doc props file: $new_file", ] );
	my	$dom = XML::LibXML->load_xml( location => $new_file );
	$self->_set_creator( ($dom->getElementsByTagName( 'dc:creator' ))[0]->textContent() );
	$self->_set_modified_by( ($dom->getElementsByTagName( 'cp:lastModifiedBy' ))[0]->textContent() );
	$self->_set_date_created(
		#~ DateTime::Format::Flexible->parse_datetime(
			($dom->getElementsByTagName( 'dcterms:created' ))[0]->textContent(),
			#~ time_zone => 'floating'
		#~ )
	);
	$self->_set_date_modified(
		#~ DateTime::Format::Flexible->parse_datetime(
			($dom->getElementsByTagName( 'dcterms:modified' ))[0]->textContent(),
			#~ time_zone => 'floating'
		#~ )
	);
	###LogSD	$phone->talk( level => 'trace', message => [ "Current object:", $self ] );
}

sub _set_shared_worksheet_files{
	my( $self, $object_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::_set_shared_worksheet_files', );
	my	$temp_dir =$self->_get_temp_dir;
	my	@file_list = <$temp_dir/xl/*>;#/
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Building the shared worksheet files with the lookup ref:", $object_ref,
	###LogSD		"Reviewing files:", @file_list ] );
	my $file_lookup;
	for my $file ( @file_list ){
		if( $file =~ /xl\/([^\.]*)\.xml$/ ){
			$file_lookup->{$1} = $file;
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"File lookup list: ", $file_lookup], );
	
	my	$name_space = $self->get_log_space;
	for my $file ( keys %$object_ref ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"checking the file class: $file",], );
		if( $file eq 'worksheet' ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Storing the worksheet superclass: ", $object_ref->{worksheet}->{superclasses}], );
			my $method = $object_ref->{$file}->{store};
			$self->$method( $object_ref->{$file}->{superclasses} );
		}elsif( exists $file_lookup->{$file} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Attempting to load the file: ${file}\.xml", ], );
			my @args = ( file_name => $file_lookup->{$file} );
			push @args, ( package => $object_ref->{$file}->{package} ) if exists $object_ref->{$file}->{package};
			push @args, ( superclasses => $object_ref->{$file}->{superclasses} ) if exists $object_ref->{$file}->{superclasses};
			for my $attribute ( @{$object_ref->{$file}->{attributes}} ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Building attribute: $attribute", ], );
				my $method = 'get_' . $attribute;
				push @args, $attribute, $self->$method;
			}
			my $role_ref;
			for my $role ( @{$object_ref->{$file}->{add_roles_in_sequence}} ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"collecting the role for: $role", ], );
				my $method = 'get_' . $role;
				push @$role_ref, $self->$method;
			}
			push @args, ( add_roles_in_sequence => $role_ref ) if $role_ref;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Final args for building the instance:", @args ], );
			my $method = $object_ref->{$file}->{store};
			#~ print "Running: $method\n" . Dumper( @args );
			$self->$method( build_instance( @args ) );
		}else{
			$self->set_error( "No file to load into the object: $file" );
		}
	}
	#~ $self->_set_current_worksheet_position( -1 );
	return 1;
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD					name_space 	=> $self->get_log_space .  '::Workbook::DEMOLISH', );
	if( $self->_has_calc_chain_file ){
		#~ print "closing calcChain.xml\n";
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD			"Clearing the calcChain.xml file" ] );
		$self->_demolish_calc_chain;
	}
	if( $self->_has_shared_strings_file ){
		my $instance = $self->_get_shared_strings_instance;
		#~ print "closing sharedStrings.xml\n" . Dumper( $instance );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD			"Clearing the sharedStrings.xml file" ] );
		if( $instance ){
			$self->_demolish_shared_strings;
		}else{
			$self->_clear_shared_strings;
			$instance = undef;
		}
	}else{
		confess "No shared strings instance found";
	}
	if( $self->_has_styles_file ){
		my $instance = $self->_get_styles_instance;
		#~ print "closing styles.xml\n" . Dumper( $instance );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD			"Clearing the styles.xml file" ] );
		if( $instance ){
			$self->_demolish_styles;
		}else{
			$self->_clear_shared_strings;
			$instance = undef;
		}
	}else{
		confess "No styles instance found";
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Clearing the Temporary Directory" ] );
	$self->_clear_temp_dir;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML - Read xlsx spreadsheet files with LibXML

=head1 SYNOPSIS

The following uses the 'TestBook.xlsx' file found in the t/test_files/ folder

	#!/usr/bin/env perl
	use strict;
	use warnings;
	use Spreadsheet::XLSX::Reader::LibXML;

	my $parser   = Spreadsheet::XLSX::Reader::LibXML->new();
	my $workbook = $parser->parse( 'TestBook.xlsx' );

	if ( !defined $workbook ) {
		die $parser->error(), "\n";
	}

	for my $worksheet ( $workbook->worksheets() ) {

		my ( $row_min, $row_max ) = $worksheet->row_range();
		my ( $col_min, $col_max ) = $worksheet->col_range();

		for my $row ( $row_min .. $row_max ) {
			for my $col ( $col_min .. $col_max ) {

				my $cell = $worksheet->get_cell( $row, $col );
				next unless $cell;

				print "Row, Col    = ($row, $col)\n";
				print "Value       = ", $cell->value(),       "\n";
				print "Unformatted = ", $cell->unformatted(), "\n";
				print "\n";
			}
		}
		last;# In order not to read all sheets
	}

	###########################
	# SYNOPSIS Screen Output
	# 01: Row, Col    = (0, 0)
	# 02: Value       = Category
	# 03: Unformatted = Category
	# 04: 
	# 05: Row, Col    = (0, 1)
	# 06: Value       = Total
	# 07: Unformatted = Total
	# 08: 
	# 09: Row, Col    = (0, 2)
	# 10: Value       = Date
	# 11: Unformatted = Date
	# 12: 
	# 13: Row, Col    = (1, 0)
	# 14: Value       = Red
	# 16: Unformatted = Red
	# 17: 
	# 18: Row, Col    = (1, 1)
	# 19: Value       = 5
	# 20: Unformatted = 5
	# 21: 
	# 22: Row, Col    = (1, 2)
	# 23: Value       = 2017-2-14 #(shows as 2/14/2017 in the sheet)
	# 24: Unformatted = 41318
	# 25: 
	# More intermediate rows ... 
	# 82: 
	# 83: Row, Col    = (6, 2)
	# 84: Value       = 2016-2-6 #(shows as 2/6/2016 in the sheet)
	# 85: Unformatted = 40944
	###########################

=head1 DESCRIPTION

This is another module for parsing Excel 2007+ workbooks.  The goal of this package is 
three fold.  First, as close as possible produce the same output as is visible in an 
excel spreadsheet with exposure to underlying settings from Excel.  Second, adhere as 
close as is reasonable to the L<Spreadsheet::ParseExcel> API (where it doesn't conflict 
with the first objective) so that less work would be needed to integrate ParseExcel and 
this package.  Third, to provide an XLSX sheet parser that is built on L<XML::LibXML>.  
The other two primary options for XLSX parsing on CPAN use either a one-off XML parser 
(L<Spreadsheet::XLSX>) or L<XML::Twig> (L<Spreadsheet::ParseXLSX>).  In general if 
either of them already work for you without issue then there is no reason to change to 
this package.  I personally found some bugs and functionality boundaries in both that I 
wanted to improve and by the time I had educated myself enough to make improvement 
suggestions including root causing the bugs to either the XML parser or the reader logic 
I had written this.

In the process of learning and building I also wrote some additional features for 
this parser that are not found in the L<Spreadsheet::ParseExcel> package.  For instance 
in the L<SYNOPSIS|/SYNOPSIS> the '$parser' and the '$workbook' are actually the same class.  
You could combine both steps by calling new with the 'file_name' attribute called out.  
Afterward it is still possible to call ->error on the instance.  Another improvement 
(From my perspective) is date handling.  This package allows for a simple pluggable custom 
output format that is more flexible than other options as well as handling dates older than 
1-January-1900.  I leveraged coercions from L<Type::Tiny> to do this but anything that 
follows that general format will work here.  Additionally, this is a L<Moose> based package.  
As such it is designed to be (fairly) extensible by writing roles and adding them to this 
package rather than requiring that you extend the package to some new branch.  Read the full 
documentation for all opportunities!

In the realm of extensibility, L<XML::LibXML> has multiple ways to read an XML file but this 
release only has an L<XML::LibXML::Reader> parser option.  Future iterations could include a 
DOM parser option.  Additionally this package does not (yet) provide the same access to the 
formatting elements provided in L<Spreadsheet::ParseExcel>.  That is on the longish and 
incomplete TODO list.

The package operates on the workbook with three primary tiers of classes.  All other classes 
in this package are for architectual extensibility.

=over

---> Workbook level (This class)

=over

---> L<Worksheet level|Spreadsheet::XLSX::Reader::LibXML::Worksheet>

=over

---> L<Cell level|Spreadsheet::XLSX::Reader::LibXML::Cell> - 
L<optional|/group_return_type>

=back

=back

=back

=head2 Primary Methods

These are the primary ways to use this class.  They can be used to open an .xlsx 
workbook.  They are also ways to investigate information at the workbook level.  For 
information on how to retrieve data from the worksheets see the 
L<Worksheet|Spreadsheet::XLSX::Reader::LibXML::Worksheet> documentation.  For additional 
workbook options see the L<Attributes|/Attributes> section.  The attributes section also 
documents all the methods used to adjust the attributes of this class.

=head3 new( %attributes )

=over

B<Definition:> This is the way to instantiate an instance of this class.  It can 
accept all the L<Attributes|/Attributes>, some, or none.  If the instance is started with 
no arguments then a L<method|/parse( $file_name, $formatter )> is needed to open the xlsx file.

B<Accepts:> the L<Attributes|/Attributes>

B<Returns:> An instance of this class

=back

=head3 parse( $file_name, $formatter )

=over

B<Definition:> This is a convenience method to match the L<Spreadsheet::ParseExcel> equivalent.  
It only works if the L<file_name|/file_name> attribute was not set with ->new.  It 
is one way to set the L<file_name|/file_name> and L<default_format_list|/default_format_list>

B<Accepts:>

	$file_name = of a valid xlsx file (required)
	$formatter = see the 'default_format_list' attribute for valid options (optional)

B<Returns:> itself when passing with the xlsx file loaded or undef for failure

=back

=head3 worksheet( $name )

=over

B<Definition:> This method will return an  object to read values in the worksheet.  
If no value is passed to $name then the 'next' worksheet in physical order is 
returned. I<'next' will NOT wrap>

B<Accepts:> the $name string representing the worksheet object you want to open

B<Returns:> a L<Worksheet|Spreadsheet::XLSX::Reader::LibXML::Worksheet> object with the 
ability to read the worksheet of that name.  Or in 'next' mode it returns undef if 
past the last sheet

B<Example:> using the implied 'next' worksheet;

	while( my $worksheet = $workbook->worksheet ){
		print "Reading: " . $worksheet->name . "\n";
		# get the data needed from this worksheet
	}

=back

=head3 start_at_the_beginning

=over

B<Definition:> This restarts the 'next' worksheet at the first worksheet

B<Accepts:>nothing

B<Returns:> nothing

=back

=head3 worksheets

=over

B<Definition:> This method will return all the worksheets in the workbook as an array. 
I<Not an array ref>.

B<Accepts:>nothing

B<Returns:> an array of L<Worksheet|Spreadsheet::XLSX::Reader::LibXML::Worksheet> 
objects with all the available worksheets in the array

=back

=head3 worksheet_name( $Int )

=over

B<Definition:> This method returns the worksheet name for a given physical position 
in the worksheet from left to right. It counts from zero even if the workbook is in 
'count_from_one' mode.

B<Accepts:>integers

B<Returns:> the worksheet name

B<Example:> To return only worksheet positions 2 through 4

	for $x (2..4){
		my $worksheet = $workbook->worksheet( $workbook->worksheet_name( $x ) );
		# Read the worksheet here
	}

=back

=head3 worksheet_names

=over

B<Definition:> This method returns an array ref of the worksheet names in the 
workbook.

B<Accepts:>nothing

B<Returns:> an array ref

B<Example:> Another way to parse a workbook without building all the sheets at 
once is;

	for $sheet_name ( @{$workbook->worksheet_names} ){
		my $worksheet = $workbook->worksheet( $sheet_name );
		# Read the worksheet here
	}

=back

=head3 number_of_sheets

=over

B<Definition:> This method returns the count of worksheets in the workbook

B<Accepts:>nothing

B<Returns:> an integer

=back

=head3 error

=over

B<Definition:> This returns the most recent error message logged by the package.  
This method is mostly relevant when an unexpected result is returned by some other 
method.

B<Accepts:>nothing

B<Returns:> an error string.

=back

=head3 get_epoch_year

=over

B<Definition:> This returns the epoch year defined by the worsheet.

B<Accepts:>nothing

B<Returns:> 1900 (= windows) or 1904 (= 1904)

=back

=head3 parse_excel_format_string( $format_string )

=over

B<Definition:> This returns a L<Type::Tiny> object with built in chained coercions 
to turn Excel Julian Dates into date strings.

B<Accepts:> a custom $format_string complying with Excel definitions 

B<Returns:> a L<Type::Tiny> object

=back

=head2 Attributes

Data passed to new when creating an instance (parser).  For modification of 
these attributes see the listed 'attribute methods'. For more information on 
attributes see L<Moose::Manual::Attributes>.

=head3 error_inst

=over

B<Definition:> This attribute holds an 'error' object instance.  It should have 
several methods for managing errors.  Currently no error codes or error translation 
options are available but this should make implementation of that easier.

B<Default:> a L<Spreadsheet::XLSX::Reader::LibXML::Error> instance with the 
attributes set as;
	
	( should_warn => 0 )

B<Range:> The minimum list of methods to implement for your own instance is;

	error set_error clear_error set_warnings if_warn

B<attribute methods> Methods provided to adjust this attribute

=over

=B<get_error_inst>

=over

B<Definition:> returns this instance

=back

B<error>

=over

B<Definition:> Used to get the most recently logged error

=back

B<set_error>

=over

B<Definition:> used to set a new error string

=back

B<clear_error>

=over

B<Definition:> used to clear the current error string in this attribute

=back

B<set_warnings>

=over

B<Definition:> used to turn on or off real time warnings when errors are set

=back

B<if_warn>

=over

B<Definition:> a method mostly used to extend this package and see if warnings 
should be emitted.

=back
		
=back

=back

=head3 file_name

=over

B<Definition:> This attribute holds the full file name and path for the 
xlsx file to be parsed.

B<Default> no default - this must be provided to read a file

B<Range> any unincrypted xlsx file that can be opened in Microsoft Excel

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_file_name>

=over

B<Definition:> change the set file name (this will reboot the workbook instance)

=back

B<has_file_name>

=over

B<Definition:> this is fundamentally a way to see if the workbook loaded correctly

=back

=back

=back

=head3 file_creator

=over

B<Definition:> This holds the information stored in the Excel Metadata 
for who created the file originally.  B<You shouldn't set this attribute 
yourself.>

B<Default> the value from the file

B<Range> A string

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<creator>

=over

B<Definition:> returns the name of the file creator

=back

=back

=back

=head3 file_date_created

=over

B<Definition:> This holds the created date in the Excel Metadata 
for when the file was first built.  B<You shouldn't set this attribute 
yourself.>

B<Default> the value from the file

B<Range> A timestamp string (ISO ish)

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<date_created>

=over

B<Definition:> returns the date the file was created

=back

=back

=back

=head3 file_modified_by

=over

B<Definition:> This holds the information stored in the Excel Metadata 
for who modified the file last.  B<You shouldn't set this attribute 
yourself.>

B<Default> the value from the file

B<Range> A string

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<modified_by>

=over

B<Definition:> returns the user name of the person who last modified the file

=back

=back

=back

=head3 file_date_modified

=over

B<Definition:> This holds the last modified date in the Excel Metadata 
for when the file was last changed.  B<You shouldn't set this attribute 
yourself.>

B<Default> the value from the file

B<Range> A timestamp string (ISO ish)

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<date_modified>

=over

B<Definition:> returns the date when the file was last modified

=back

=back

=back

=head3 sheet_parser

=over

B<Definition:> This sets the way the .xlsx file is parsed.  For now the only 
choice is 'reader'.

B<Default> 'reader'

B<Range> 'reader'

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_parser_type>

=over

B<Definition:> the way to change the parser type

=back

B<get_parser_type>

=over

B<Definition:> returns the currently set parser type

=back

=back

=back

=head3 count_from_zero

=over

B<Definition:> Excel spreadsheets count from 1.  L<Spreadsheet::ParseExcel> 
counts from zero.  This allows you to choose either way.

B<Default> 1

B<Range> 1 = counting from zero like Spreadsheet::ParseExcel, 
0 = Counting from 1 lke Excel

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<counting_from_zero>

=over

B<Definition:> a way to check the current attribute setting

=back

B<set_count_from_zero>

=over

B<Definition:> a way to change the current attribute setting

=back

=back

=back

=head3 file_boundary_flags

=over

B<Definition:> When you request data past the end of a row or past the bottom 
of the data this package can return 'EOR' or 'EOF' to indicate that state.  
This is especially helpful in 'while' loops.  The other option is to return 
'undef'.  This is problematic if some cells in your table are empty which also 
returns undef.

B<Default> 1

B<Range> 1 = return 'EOR' or 'EOF' flags as appropriate,
0 = return undef when requesting a position that is out of bounds

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<boundary_flag_setting>

=over

B<Definition:> a way to check the current attribute setting

=back

B<change_boundary_flag>

=over

B<Definition:> a way to change the current attribute setting

=back

=back

=back

=head3 empty_is_end

=over

B<Definition:> The excel convention is to read the table left to right and top 
to bottom.  Some tables have uneven columns from row to row.  This allows the 
several methods that take 'next' values to wrap after the last element with data 
rather than going to the max column.

B<Default> 0

B<Range> 1 = treat all columns short of the max column for the sheet as being in 
the table, 0 = end each row after the last cell with data rather than going to the 
max sheet column

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<is_empty_the_end>

=over

B<Definition:> a way to check the current attribute setting

=back

B<set_empty_is_end>

=over

B<Definition:> a way to set the current attribute setting

=back

=back

=back

=head3 from_the_edge

=over

B<Definition:> Some data tables start in the top left corner.  Others do not.  I 
don't reccomend that practice but when aquiring data in the wild it is often good 
to adapt.  This attribute sets whether the file reads from the top left edge or from 
the top row with data and starting from the leftmost column with data.

B<Default> 1

B<Range> 1 = treat the top left corner of the sheet even if there is no data in 
the top row or leftmost column, 0 = Set the minimum row and minimum columns to be 
the first row and first column with data

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_from_the_edge>

=over

B<Definition:> a way to set the current attribute setting

=back

=back

=back

=head3 default_format_list

=over

B<Definition:> This is a departure from L<Spreadsheet::ParseExcel> for two reasons.  
First, it doesn't use the same modules.  Second, this accepts a role with two methods 
where ParseExcel accepts an object instance.

B<Default> Spreadsheet::XLSX::Reader::LibXML::FmtDefault

B<Range> a L<Moose> role with the methods 'get_defined_excel_format' and 
'change_output_encoding' it should be noted that libxml2 which is the underlying code 
for L<XML::LibXML> allways attempts to get the data into perl friendly strings.  That 
means this should only tweak the data on the way out and does not affect the data on the 
way in.

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_default_format_list>

=over

B<Definition:> a way to check the current attribute setting

=back

B<set_default_format_list>

=over

B<Definition:> a way to set the current attribute setting

=back

=back

=back

=head3 format_string_parser

=over

B<Definition:> This is the interpreter that turns the excel into a L<Type::Tiny> coercion.  
If you don't like the output or the method you can write your own Moose Role and add it here.

B<Default> Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings

B<Range> a L<Moose> role with the method 'parse_excel_format_string'

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_format_string_parser>

=over

B<Definition:> a way to check the current attribute setting

=back

B<set_format_string_parser>

=over

B<Definition:> a way to set the current attribute setting

=back

=back

=back

=head3 group_return_type

=over

B<Definition:> Traditionally ParseExcel returns a cell object with lots of methods 
to reveal information about the cell.  In reality this is probably not used very much 
so in the interest of simplifying you can get a cell object instance set to the cell 
information.  Or you can just get the raw value in the cell or you can get the cell value 
formatted either the way the sheet specified or the way you specify.  See the 
'custom_formats' attribute for the L<Spreadsheet::XLSX::Reader::LibXML::Worksheet> class 
to insert custom targeted formats for use with the parser.  All empty cells return undef 
no matter what.

B<Default> instance

B<Range> instance = returns a populated L<Spreadsheet::XLSX::Reader::LibXML::Cell> instance,
unformatted = returns the raw value of the cell with no modifications, value = returns just 
the formatted value stored in the excel cell

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<get_group_return_type>

=over

B<Definition:> a way to check the current attribute setting

=back

B<set_group_return_type>

=over

B<Definition:> a way to set the current attribute setting

=back

=back

=back

=head1 BUILD / INSTALL from Source

B<1.> Ensure that you have the libxml2 B<and libxml2-devel> libraries installed using 
your favorite package installer

L<http://xmlsoft.org/>
	
B<2.> Download a compressed file with the code from your favorite source
	
B<3.> Extract the code from the compressed file.  If you are using tar this should work:

        tar -zxvf Spreadsheet-XLSX-Reader-LibXML-v0.xx.tar.gz

B<4.> Change (cd) into the extracted directory

B<5.> Run the following

=over

(For Windows find what version of make was used to compile your perl)

	perl  -V:make
	
(for Windows below substitute the correct make function (s/make/dmake/g)?)
	
=back

	>perl Makefile.PL

	>make

	>make test

	>make install # As sudo/root

	>make clean

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Build L<Alien::LibXML::Devel> to load the libxml2-devel libraries from source and 
require that and L<Alien::LibXML> in the build file. So all needed requirements for L<XML::LibXML> 
are met

=over

Both libxml2 and libxml2-devel libraries are required for XML::LibXML

=back

B<2.> Add a pivot table reader (Not just read the values from the sheet)

B<3.> Add calc chain methods

B<4.> Add more exposure to workbook formatting methods

B<5.> Build a DOM parser alternative for the sheets

=over

(Theoretically faster than the reader but uses more memory)

=back

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

L<perl 5.010|perl/5.10.0>

L<version>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<Archive::Zip>

L<OLE::Storage_Lite>

L<File::Temp>

L<Type::Tiny> - 0.046

L<MooseX::ShortCut::BuildInstance> - 1.026

L<Carp>

L<XML::LibXML>

L<Clone>

L<DateTimeX::Format::Excel>

L<DateTime::Format::Flexible>

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