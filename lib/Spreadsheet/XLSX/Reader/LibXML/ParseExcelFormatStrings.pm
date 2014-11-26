package Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings;
use version; our $VERSION = qv('v0.12.4');

use 5.010;
use Moose::Role;
requires qw(
	get_log_space				_add_integer_separator
	_continuous_fraction		change_output_encoding
	get_excel_region
);
use Types::Standard qw(
		Int						Str				
		Maybe					Num					
		HashRef					ArrayRef
		CodeRef					Object			
		ConsumerOf				InstanceOf			
		HasMethods				Bool
		is_Object
    );
use Carp qw( confess );
use	Type::Coercion;
use	Type::Tiny;
use DateTimeX::Format::Excel 0.012;
use DateTime::Format::Flexible;
use Clone 'clone';
use lib	'../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Types 0.012 qw(
		PositiveNum				NegativeNum
		ZeroOrUndef				NotNegativeNum
		Excel_number_0
	);

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9

my	$coercion_index		= 0;
my	@type_list			= ( PositiveNum, NegativeNum, ZeroOrUndef, Str );
my	$last_date_cldr		= 'yyyy-m-d';
my	$last_duration		= 0;
my	$last_sub_seconds	= 0;
my	$last_format_rem	= 0;
my	$duration_order		={ h => 'm', m =>'s', s =>'0' };
my	$number_build_dispatch ={
		all =>[qw(
			_convert_negative
			_divide_by_thousands
			_convert_to_percent
			_split_decimal_integer
			_move_decimal_point
			_build_fraction
			_round_decimal
			_add_commas
			_pad_exponent
		)],
		scientific =>[qw(
			_convert_negative
			_split_decimal_integer
			_move_decimal_point
			_round_decimal
			_add_commas
			_pad_exponent
		)],
		percent =>[qw(
			_convert_negative
			_convert_to_percent
			_split_decimal_integer
			_round_decimal
			_add_commas
		)],
		fraction =>[qw(
			_convert_negative
			_split_decimal_integer
			_build_fraction
			_add_commas
		)],
		integer =>[qw(
			_convert_negative
			_divide_by_thousands
			_split_decimal_integer
			_round_decimal
			_add_commas
		)],
		decimal =>[qw(
			_convert_negative
			_divide_by_thousands
			_split_decimal_integer
			_round_decimal
			_add_commas
		)],
	};

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has	epoch_year =>(
		isa		=> Int,
		reader	=> 'get_epoch_year',
		default	=> 1900,
	);
	
has	cache_formats =>(
		isa		=> Bool,
		reader	=> 'get_cache_behavior',
		default	=> 1,
	);
	
has	datetime_dates =>(
		isa		=> Bool,
		reader	=> 'get_date_behavior',
		writer	=> 'set_date_behavior',
		default	=> 0,
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9
 
sub parse_excel_format_string{# Currently only handles dates and times
	my( $self, $format_strings, $coercion_name ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space .  '::parse_excel_format_string', );
	if( !defined $format_strings ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Nothing passed to convert",] );
		return Excel_number_0;
	}
	$format_strings =~ s/\\//g;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"parsing the custom excel format string: $format_strings",] );
	my $conversion_type = 'number';
	# Check the cache
	my	$cache_key;
	if( $self->get_cache_behavior ){
		$cache_key	= $format_strings; # TODO fix the non-hashkey character issues;
		if( $self->has_cached_format( $cache_key ) ){
			###LogSD		$phone->talk( level => 'debug', message => [
			###LogSD			"Format already built - returning stored value for: $cache_key", ] );
			return $self->get_cached_format( $cache_key );
		}else{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Building new format for key: $cache_key", ] );
		}
	}
	
	# Split into the four sections positive, negative, zero, and text
		$format_strings =~ s/General/\@/ig;# Change General to text input
	my	@format_string_list = split /;/, $format_strings;
	my	$last_is_text = ( $format_string_list[-1] =~ /\@/ ) ? 1 : 0 ;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Is the last position text: $last_is_text",	] );
	# Make sure the full range of number inputs are sent down the right path;
	my	@used_type_list = @{\@type_list};
		$used_type_list[0] =
			( scalar( @format_string_list ) - $last_is_text == 1 ) ? Maybe[Num] :
			( scalar( @format_string_list ) - $last_is_text == 2 ) ? Maybe[NotNegativeNum] : $type_list[0] ;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Now operating on each format string", @format_string_list,
	###LogSD		'..with used type list:', map{ $_->name } @used_type_list,	] );
	my	$format_position = 0;
	my	@coercion_list;
	my	$action_type;
	my	$is_date = 0;
	my	$date_text = 0;
	for my $format_string ( @format_string_list ){
		$format_string =~ s/_.//g;# no character justification to other rows
		$format_string =~ s/\*//g;# Remove the repeat character listing (not supported here)
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Building format for: $format_string", ] );
		
		# Pull out all the straight through stuff
		my @deconstructed_list;
		my $x = 0;
		while( defined $format_string and my @result = $format_string =~
					/^(													# Collect any formatting stuff first
						(AM\/PM|										# Date 12 hr flag
						A\/P|											# Another date 12 hr flag
						\[hh?\]|										# Elapsed hours
						\[mm\]|											# Elapsed minutes
						\[ss\]|											# Elapsed seconds
						[dmyhms]+)|										# DateTime chunks
						([0-9#\?]+[,\-\_]?[#0\?]*,*|					# Number string
						\.|												# Split integers from decimals
						[Ee][+\-]|										# Exponential notiation
						%)|												# Percentage
						(\@)											# Text input
					)?(													# Finish collecting format actions
					(\"[^\"]*\")|										# Anything in quotes just passes through
					(\[[^\]]*\])|										# Anything in brackets needs modification
					[\(\)\$\-\+\/\:\!\^\&\'\~\{\}\<\>\=\s]|				# All the pass through characters
					\,\s												# comma space for verbal separation
					)?(.*)/x											){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Now processing: $format_string", '..with result:', @result ] );
			my	$pre_action		= $1;
			my	$date			= $2;
			my	$number			= $3;
			my	$text			= $4;
			my	$fixed_value	= $5;
				$format_string	= $8;
			if( $fixed_value ){
				if( $fixed_value =~ /\[\$([^\-\]]*)\-?\d*\]/ ){# removed the localized element of fixed values
					$fixed_value = $1;
				}elsif( $fixed_value =~ /\[[^hms]*\]/ ){# Remove all color and conditionals as they will not be used
					$fixed_value = undef;
				}elsif( $fixed_value =~ /\"\-\"/ and $format_string ){# remove decimal justification for zero bars
					###LogSD	$phone->talk( level => 'trace', message => [
					###LogSD		"Initial format string: $format_string", ] );
					$format_string =~ s/^(\?+)//;
					###LogSD	$phone->talk( level => 'trace', message => [
					###LogSD		"updated format string: $format_string", ] );
				}
			}
			if( defined $pre_action ){
				my	$current_action =
						( $date ) ? 'DATE' :
						( defined $number ) ? 'NUMBER' :
						( $text ) ? 'TEXT' : 'BAD' ;
					$is_date = 1 if $date;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Current action from -$pre_action- is: $current_action" ] );
				if( $action_type and $current_action and ($current_action ne $action_type) ){
					###LogSD	$phone->talk( level => 'info', message => [
					###LogSD		"General action type: $action_type",,
					###LogSD		"is failing current action: $current_action", ] );
					my $fail = 1;
					if( $action_type eq 'DATE' ){
						$conversion_type = 'date';
						###LogSD	$phone->talk( level => 'info', message => [
						###LogSD		"Checking the date mishmash", ] );
						if( $current_action eq 'NUMBER' ){
							###LogSD	$phone->talk( level => 'info', message => [
							###LogSD		"Special case of number following action", ] );
							if(	( $pre_action =~ /^\.$/ and $format_string =~ /^0+/				) or
								( $pre_action =~ /^0+$/ and $deconstructed_list[-1]->[0] =~ /^\.$/	)	){
								$current_action = 'DATE';
								$fail = 0;
							}
						}elsif( $pre_action eq '@' ){
							###LogSD	$phone->talk( level => 'info', message => [
							###LogSD		"Excel conversion of pre-epoch datestring pass through highjacked here", ] );
							$current_action = 'DATESTRING';
							$fail = 0;
						}
					}elsif( $action_type eq 'NUMBER' ){
						###LogSD	$phone->talk( level => 'info', message => [
						###LogSD		"Checking for possible text in a number field for a pass throug", ] );
						if( $current_action eq 'TEXT' ){
							###LogSD	$phone->talk( level => 'info', message => [
							###LogSD		"Special case of text following a number", ] );
							$fail = 0;
						}
					}
					if( $fail ){
						confess "Bad combination of actions in this format string: $format_strings - $action_type - $current_action";
					}
				}
				$action_type = $current_action if $current_action;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		(($pre_action) ? "First action resolved to: $pre_action" : undef),
				###LogSD		(($fixed_value) ? "Extracted fixed value: $fixed_value" : undef),
				###LogSD		(($format_string) ? "Remaining string: $format_string" : undef),
				###LogSD		"With updated deconstruction list:", @deconstructed_list, ] );
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Early elements unusable - remaining string: $format_string", ] );
			}
			push @deconstructed_list, [ $pre_action, $fixed_value ];
			if( $x++ == 30 ){
				confess "Regex matching failed (with an infinite loop) for excel format string: $format_string";
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		(($pre_action) ? "First action resolved to: $pre_action" : undef),
			###LogSD		(($fixed_value) ? "Extracted fixed value: $fixed_value" : undef),
			###LogSD		(($format_string) ? "Remaining string: $format_string" : undef),
			###LogSD		"With updated deconstruction list:", @deconstructed_list, ] );
			last if length( $format_string ) == 0;
		}
		push @deconstructed_list, [ $format_string, undef ] if $format_string;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Current segment type: $action_type", 
		###LogSD		"List with fixed values separated:", @deconstructed_list ] );
		my $method = '_build_' . lc($action_type);
		my $filter = ( $action_type eq 'TEXT' ) ? Str : $used_type_list[$format_position++];
		if( $action_type eq 'DATESTRING' ){
			$date_text = 1;
			$filter = Str;
		}
		push @coercion_list, $self->$method( $filter, \@deconstructed_list );
	}
	push @coercion_list, $self->_build_datestring( Str, [ [ '@', '' ] ] ) if $is_date and !$date_text;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Length of coersion list: ' . scalar( @coercion_list ),
	###LogSD		(map{ if( is_Object( $_ ) and $_->can( 'name' ) ){ $_->name }else{ $_ } } @coercion_list), ] );
	
	# Build the final format
	$coercion_name =~ s/__/_${conversion_type}_/ if $coercion_name;
	my	%args = (
			name		=> ($coercion_name // ($action_type . '_' . $coercion_index++)),
			coercion	=>	Type::Coercion->new(
								type_coercion_map => [ @coercion_list ],
							),
			#~ coerce		=> 1,
		);
	my	$final_type = Type::Tiny->new( %args );
	
	# Save the cache
	$self->set_cached_format( $cache_key => $final_type ) if $self->get_cache_behavior;
	
	return $final_type;
}
	

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9
	
has	_format_cash =>(
		isa		=> HashRef,
		traits	=> ['Hash'],
		handles =>{
			has_cached_format => 'exists',
			get_cached_format => 'get',
			set_cached_format => 'set',
		},
		default	=> sub{ {} },
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _build_text{
	my( $self, $type_filter, $list_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_text', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to process text values" ] );
	my $sprintf_string;
	my $found_string = 0;
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing text piece:", $piece ] );
		if( !$found_string and defined $piece->[0] ){
			$sprintf_string .= '%s';
			$found_string = 1;
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	my	$return_sub = sub{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Updated Input: $_[0]" ] );
			return sprintf( $sprintf_string, $_[0] );
		};
	return( Str, $return_sub );
}

sub _build_date{
	my( $self, $type_filter, $list_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_date', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to process date values" ] );
	
	my ( $cldr_string, $format_remainder );
	my	$is_duration = 0;
	my	$sub_seconds = 0;
	if( !$self->get_date_behavior ){
		# Process once to build the cldr string
		my $prior_duration;
		for my $piece ( @$list_ref ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"processing date piece:", $piece ] );
			if( defined $piece->[0] ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Manageing the cldr part: " . $piece->[0] ] );
				if( $piece->[0] =~ /\[(.+)\]/ ){
					###LogSD	$phone->talk( level => 'debug', message =>[ "Possible duration" ] );
					(my $initial,) = split //, $1;
					my $length = length( $1 );
					$is_duration = [ $initial, 0, [ $piece->[1] ], [ $length ] ];
					if( $is_duration->[0] =~ /[hms]/ ){
						$piece->[0] = '';
						$piece->[1] = '';
						$prior_duration = 	$is_duration->[0];
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"found a duration piece:", $is_duration,
						###LogSD		"with prior duration: $prior_duration"		] );
					}else{
						confess "Bad duration element found: $is_duration->[0]";
					} 
				}elsif( ref( $is_duration ) eq 'ARRAY' ){
					###LogSD	$phone->talk( level => 'debug', message =>[ "adding to duration", $piece ] );
					my	$next_duration = $duration_order->{$prior_duration};
					if( $piece->[0] eq '.' ){
						push @{$is_duration->[2]}, $piece->[0];
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"found a period" ] );
					}elsif( $piece->[0] =~ /$next_duration/ ){
						my $length = length( $piece->[0] );
						$is_duration->[1]++;
						push @{$is_duration->[2]}, $piece->[1] if $piece->[1];
						push @{$is_duration->[3]}, $length;
						($prior_duration,) = split //, $piece->[0];
						if( $piece->[0] =~ /^0+$/ ){
							$piece->[0] =~ s/0/S/g;
							$sub_seconds = $piece->[0];
							###LogSD	$phone->talk( level => 'debug', message => [
							###LogSD		"found a subseconds format piece: $sub_seconds" ] );
						}
						$piece->[0] = '';
						$piece->[1] = '';
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Current duration:", $is_duration,
						###LogSD		"with prior duration: $prior_duration"	 ] );
					}else{
						confess "Bad duration element found: $piece->[0]";
					} 
				}elsif( $piece->[0] =~ /m/ ){
					###LogSD	$phone->talk( level => 'debug', message =>[ "Minutes or Months" ] );
					if( ($cldr_string and $cldr_string =~ /:'?$/) or ($piece->[1] and $piece->[1] eq ':') ){
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Found minutes - leave them alone" ] );
					}else{
						$piece->[0] =~ s/m/L/g;
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Converting to cldr stand alone months (m->L)" ] );
					}
				}elsif( $piece->[0] =~ /h/ ){
					$piece->[0] =~ s/h/H/g;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Converting 12 hour clock to 24 hour clock" ] );
				}elsif( $piece->[0] =~ /AM?\/PM?/i ){
					$cldr_string =~ s/H/h/g;
					$piece->[0] = 'a';
					###LogSD	$phone->talk( level => 'debug', message =>[ "Set 12 hour clock and AM/PM" ] );
				}elsif( $piece->[0] =~ /d{3,5}/ ){
					$piece->[0] =~ s/d/E/g;
					###LogSD	$phone->talk( level => 'debug', message =>[ "Found a weekday request" ] );
				}elsif( !$sub_seconds and $piece->[0] =~ /[\.]/){#
					$piece->[0] = "'.'";
					#~ $piece->[0] = "':'";
					$sub_seconds = 1;
					###LogSD	$phone->talk( level => 'debug', message =>[ "Starting sub seconds" ] );
				}elsif( $sub_seconds eq '1' ){
					###LogSD	$phone->talk( level => 'debug', message =>[ "Formatting sub seconds" ] );
					if( $piece->[0] =~ /^0+$/ ){
						$piece->[0] =~ s/0/S/g;
						$sub_seconds = $piece->[0];
						$piece->[0] = '';
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"found a subseconds format piece: $sub_seconds" ] );
					}else{
						confess "Bad sub-seconds element after [$cldr_string] found: $piece->[0]";
					}
				}
				if( $sub_seconds and $sub_seconds ne '1' ){
					$format_remainder .= $piece->[0];
				}else{
					$cldr_string .= $piece->[0];
				}
			}
			if( $piece->[1] ){
				if( $sub_seconds and $sub_seconds ne '1' ){
					$format_remainder .= $piece->[1];
				}else{
					$cldr_string .= $piece->[1];
				}
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		(($cldr_string) ? "Updated CLDR string: $cldr_string" : undef),
			###LogSD		(($format_remainder) ? "Updated format remainder: $format_remainder" : undef),
			###LogSD		(($is_duration) ? ('Duration ref:', $is_duration) : undef)			] );
		}
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Updated CLDR string: $cldr_string",
		###LogSD		(($is_duration) ? ('...and duration:', $is_duration) : undef )	] );
		$last_date_cldr 	= $cldr_string;
		$last_duration		= $is_duration;
		$last_sub_seconds	= $sub_seconds;
		$last_format_rem	= $format_remainder;
	}
	my	@args_list = ( $self->get_epoch_year == 1904 ) ? ( system_type => 'apple_excel' ) : ();
	my	$converter = DateTimeX::Format::Excel->new( @args_list );
	my	$conversion_sub = sub{ 
			my	$num = $_[0];
			if( !defined $num ){
				return undef;
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing date number: $num",
			###LogSD		'..with duration:', $is_duration,
			###LogSD		"..and sub-seconds: $sub_seconds",
			###LogSD		(($format_remainder) ? "..and format_remainder: $format_remainder" : undef) ] );
			my	$dt = $converter->parse_datetime( $num );
			my $return_string;
			my $calc_sub_secs;
			if( $is_duration ){
				my	$di = $dt->subtract_datetime_absolute( $converter->_get_epoch_start );
				if( $self->get_date_behavior ){
					return $di;
				}
				my	$sign = DateTime->compare_ignore_floating( $dt, $converter->_get_epoch_start );
				$return_string = ( $sign == -1 ) ? '-' : '' ;
				my $key = $is_duration->[0];
				my $delta_seconds	= $di->seconds;
				my $delta_nanosecs	= $di->nanoseconds;
				$return_string .= $self->_build_duration( $is_duration, $delta_seconds, $delta_nanosecs );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Duration return string: $return_string" ] );
			}else{
				if( $self->get_date_behavior ){
					return $dt;
				}
				if( $sub_seconds ){
					$calc_sub_secs = $dt->format_cldr( $sub_seconds );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Processing sub-seconds: $calc_sub_secs" ] );
					if( "0.$calc_sub_secs" >= 0.5 ){
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Rounding seconds back down" ] );
						$dt->subtract( seconds => 1 );
					}
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Converting it with CLDR string: $cldr_string" ] );
				$return_string .= $dt->format_cldr( $cldr_string );
				if( $sub_seconds and $sub_seconds ne '1' ){
					$return_string .= $calc_sub_secs;
				}
				$return_string .= $dt->format_cldr( $format_remainder ) if $format_remainder;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"returning: $return_string" ] );
			}
			return $return_string;
		};
	return( $type_filter, $conversion_sub );
}

sub _build_datestring{
	my( $self, $type_filter, $list_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_datestring', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to process date strings" ] );
	
	my ( $cldr_string, $format_remainder );
	#~ my	$is_duration = 0;
	#~ if( !$self->get_date_behavior ){
		#~ $cldr_string = $last_date_cldr;
	#~ }
	my	$conversion_sub = sub{ 
			my	$date = $_[0];
			if( !$date ){
				return undef;
			}
			my $calc_sub_secs;
			if( $date =~ /(.*:\d+)\.(\d+)(.*)/ ){
				$calc_sub_secs = $2;
				$date = $1;
				$date .= $3 if $3;
				$calc_sub_secs .= 0 x (9 - length( $calc_sub_secs ));
			}
			my	$dt = 	DateTime::Format::Flexible->parse_datetime(
							$date, lang =>[ $self->get_excel_region ]
						);
			$dt->add( nanoseconds => $calc_sub_secs ) if $calc_sub_secs;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing date string: $date",
			###LogSD		"..with duration:", $last_duration,
			###LogSD		"..and sub-seconds: $last_sub_seconds",
			###LogSD		"..and stripped nanoseconds: $calc_sub_secs"		] );
			my $return_string;
			if( $last_duration ){
				my	@args_list = ( $self->get_epoch_year == 1904 ) ? ( system_type => 'apple_excel' ) : ();
				my	$converter = DateTimeX::Format::Excel->new( @args_list );
				my	$di = $dt->subtract_datetime_absolute( $converter->_get_epoch_start );
				if( $self->get_date_behavior ){
					return $di;
				}
				my	$sign = DateTime->compare_ignore_floating( $dt, $converter->_get_epoch_start );
				$return_string = ( $sign == -1 ) ? '-' : '' ;
				my $key = $last_duration->[0];
				my $delta_seconds	= $di->seconds;
				my $delta_nanosecs	= $di->nanoseconds;;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Delta seconds: $delta_seconds",
				###LogSD		(($delta_nanosecs) ? "Delta nanoseconds: $delta_nanosecs" : undef) ] );
				$return_string .= $self->_build_duration( $last_duration, $delta_seconds, $delta_nanosecs );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Duration return string: $return_string" ] );
			}else{
				if( $self->get_date_behavior ){
					return $dt;
				}
				if( $last_sub_seconds ){
					$calc_sub_secs = $dt->format_cldr( $last_sub_seconds );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Processing sub-seconds: $calc_sub_secs" ] );
					if( "0.$calc_sub_secs" >= 0.5 ){
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Rounding seconds back down" ] );
						$dt->subtract( seconds => 1 );
					}
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Converting it with CLDR string: $last_date_cldr" ] );
				$return_string .= $dt->format_cldr( $last_date_cldr );
				if( $last_sub_seconds and $last_sub_seconds ne '1' ){
					$return_string .= $calc_sub_secs;
				}
				$return_string .= $dt->format_cldr( $last_format_rem ) if $last_format_rem;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"returning: $return_string" ] );
			}
			return $return_string;
		};
	return( $type_filter, $conversion_sub );
}

sub _build_duration{
	my( $self, $duration_ref, $delta_seconds, $delta_nanosecs ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_date::_build_duration', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Building a duration string with duration ref:', $duration_ref,
	###LogSD			"With delta seconds: $delta_seconds",
	###LogSD			(($delta_nanosecs) ? "And delta nanoseconds: $delta_nanosecs" : undef) ] );
	my	$return_string;
	my	$key = $duration_ref->[0];
	my	$first = 1;
	for my $position ( 0 .. $duration_ref->[1] ){
		if( $key eq '0' ){
			my $length = length( $last_sub_seconds );
			$return_string .= '.' . sprintf( "%0.${length}f", $delta_nanosecs/1000000000);
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Return string with nanoseconds: $return_string", ] );
		}
		if( $key eq 's' ){
			$return_string .= ( $first ) ? $delta_seconds :
				sprintf "%0$duration_ref->[3]->[$position]d", $delta_seconds;
			$first = 0;
			$key = $duration_order->{$key};
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Delta seconds: $delta_seconds",
			###LogSD		"Next key to process: $key"			] );
		}
		if( $key eq 'm' ){
			my $minutes = int($delta_seconds/60);
			$delta_seconds = $delta_seconds - ($minutes*60);
			$return_string .= ( $first ) ? $minutes :
				sprintf "%0$duration_ref->[3]->[$position]d", $minutes;
			$first = 0;
			$key = $duration_order->{$key};
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Calculated minutes: $minutes",
			###LogSD		"Remaining seconds: $delta_seconds",
			###LogSD		"Next key to process: $key"			] );
		}
		if( $key eq 'h' ){
			my $hours = int($delta_seconds /(60*60));
			$delta_seconds = $delta_seconds - ($hours*60*60);
			$return_string .= ( $first ) ? $hours :
				sprintf "%0$duration_ref->[3]->[$position]d", $hours;
			$first = 0;
			$key = $duration_order->{$key};
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Calculated hours: $hours",
			###LogSD		"Remaining seconds: $delta_seconds",
			###LogSD		"Next key to process: $key"			] );
		}
		$return_string .= $duration_ref->[2]->[$position] if $duration_ref->[2]->[$position];
	}
	return $return_string;
}

sub _build_number{
	my( $self, $type_filter, $list_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Processing a number list to see how it should be converted",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref 			] );
	my ( $code_hash_ref, $number_type, );
	
	# Resolve zero replacements quickly
	if(	$type_filter->name eq 'ZeroOrUndef' and
		!$list_ref->[-1]->[0] and $list_ref->[-1]->[1] eq '"-"' ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Found a zero to bar replacement"			] );
		my $return_string;
		for my $piece ( @$list_ref ){
			$return_string .= $piece->[1];
		}
		$return_string =~ s/"\-"/\-/;
		return( $type_filter, sub{ $return_string } );
	}
	
	# Process once to determine what to do
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( defined $piece->[0] ){
			if( my @result = $piece->[0] =~ /^([0-9#\?]+)([,\-\_])?([#0\?]+)?(,+)?$/ ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Regex yielded result:", @result ] );
				my	$comma = ($2) ? $2 : undef,
				my	$comma_less = defined( $3) ? "$1$3" : $1;
				my	$comma_group = length( $3 );
				my	$divide_by_thousands = ( $4 ) ? (( $2 and $2 ne ',' ) ? $4 : "$2$4" ) : undef;#eval{ $2 . $4 }
				my	$divisor = $1 if $1 =~ /^([0-9]+)$/;
				my ( $leading_zeros, $trailinq_zeros );
				if( $comma_less =~ /^[\#\?]*(0+)$/ ){
					$leading_zeros = $1;
				}
				if( $comma_less =~ /^(0+)[\#\?]*$/ ){
					$trailinq_zeros = $1;
				}
				$code_hash_ref->{divide_by_thousands} = length( $divide_by_thousands ) if $divide_by_thousands;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"The comma less string is extracted to: $comma_less",
				###LogSD		((defined $comma_group) ? "The separator group length is: $comma_group" : undef),
				###LogSD		(($comma) ? "The separator character is: $comma" : undef),
				###LogSD		((length( $leading_zeros )) ? ".. w/leading zeros: $leading_zeros" : undef),
				###LogSD		((length( $trailinq_zeros )) ? ".. w/trailing zeros: $trailinq_zeros" : undef),
				###LogSD		(($divisor) ? "..with identified divisor: $divisor" : undef),
				###LogSD		'Initial code hash:', $code_hash_ref] );
				if( !$number_type ){
					$number_type = 'INTEGER';
					$code_hash_ref->{integer}->{leading_zeros} = length( $leading_zeros ) if length( $leading_zeros );
					$code_hash_ref->{integer}->{minimum_length} = length( $comma_less );
					if( $comma ){
						@{$code_hash_ref->{integer}}{ 'group_length', 'comma' } = ( $comma_group, $comma );
					}
					if( defined $piece->[1] ){
						if( $piece->[1] =~ /(\s+)/ ){
							$code_hash_ref->{separator} = $1;
						}elsif( $piece->[1] eq '/' ){
							$number_type = 'FRACTION';
							$code_hash_ref->{numerator}->{leading_zeros} = length( $leading_zeros ) if length( $leading_zeros );
							delete $code_hash_ref->{integer};
						}
					}
				}elsif( ($number_type eq 'INTEGER') or $number_type eq 'DECIMAL' ){
					if( $piece->[1] and $piece->[1] eq '/'){
						$number_type = 'FRACTION';
					}else{
						$number_type = 'DECIMAL';
						$code_hash_ref->{decimal}->{trailing_zeros} = length( $trailinq_zeros ) if length( $trailinq_zeros );
						$code_hash_ref->{decimal}->{max_length} = length( $comma_less );
					}
				}elsif( ($number_type eq 'SCIENTIFIC') or $number_type eq 'FRACTION' ){
					$code_hash_ref->{exponent}->{leading_zeros} = length( $leading_zeros ) if length( $leading_zeros );
					$code_hash_ref->{fraction}->{target_length} = length( $comma_less );
					if( $divisor ){
						$code_hash_ref->{fraction}->{divisor} = $divisor;
					}
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Current number type: $number_type", 'updated settings:', $code_hash_ref] );
			}elsif( $piece->[0] =~ /^((\.)|([Ee][+\-])|(%))$/ ){
				if( $2 ){
					$number_type = 'DECIMAL';
					$code_hash_ref->{separator} = $1;
				}elsif( $3 ){
					$number_type = 'SCIENTIFIC';
					$code_hash_ref->{separator} = $2;
				}else{
					$number_type = 'PERCENT';
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Number type now: $number_type" ] );
			}else{
				confess "badly formed number format passed: $piece->[0]";
			}
		}
	}
	
	my $method = '_build_' . lc( $number_type ) . '_sub';
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Resolved the number type to: $number_type",
	###LogSD		'Working with settings:', $code_hash_ref ] );
	
	# Set negative type
	if( $type_filter->name eq 'NegativeNum' ){
		$code_hash_ref->{negative_type} = 1;
	}
	my $conversion_sub = $self->$method( $type_filter, $list_ref, $code_hash_ref );
		
	return( $type_filter, $conversion_sub );
}

sub _build_integer_sub{
	my( $self, $type_filter, $list_ref, $conversion_defs ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_integer_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return integer values",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and conversion defs:', $conversion_defs	] );
	
	my $sprintf_string;
	# Process once to determine what to do
	my $found_integer = 0;
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( !$found_integer and defined $piece->[0] ){
			$sprintf_string .= '%s';
			$found_integer = 1;
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	$conversion_defs->{no_decimal} = 1;
	$conversion_defs->{sprintf_string} = $sprintf_string;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	my $dispatch_sequence = $number_build_dispatch->{decimal};
	
	my 	$conversion_sub = sub{
			my $adjusted_input = $_[0];
			if( !defined $adjusted_input or $adjusted_input eq '' ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			my	$value_definitions = clone( $conversion_defs );
				$value_definitions->{initial_value} = $adjusted_input;
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'Building scientific output with:',  $conversion_defs,
			###LogSD		'..and dispatch sequence:', $dispatch_sequence ] );
			my $built_ref = $self->_build_elements( $dispatch_sequence, $value_definitions );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Received built ref:", $built_ref ] );
			my	$return .= sprintf(
					$built_ref->{sprintf_string},
					$built_ref->{integer}->{value}
				);
			$return = $built_ref->{sign} . $return if $built_ref->{sign} and $return;
			return $return;
		};
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

sub _build_decimal_sub{
	my( $self, $type_filter, $list_ref, $conversion_defs ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_decimal_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return decimal values",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and code hash ref:', $conversion_defs ] );
	
	my $sprintf_string;
	# Process once to determine what to do
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( defined $piece->[0] ){
			if( $piece->[0] eq '.' ){
				$sprintf_string .= '.';
			}else{
				$sprintf_string .= '%s';
			}
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	$conversion_defs->{sprintf_string} = $sprintf_string;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	my $dispatch_sequence = $number_build_dispatch->{decimal};
	
	my 	$conversion_sub = sub{
			my $adjusted_input = $_[0];
			if( !defined $adjusted_input or $adjusted_input eq '' ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			my	$value_definitions = clone( $conversion_defs );
				$value_definitions->{initial_value} = $adjusted_input;
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'Building scientific output with:',  $conversion_defs,
			###LogSD		'..and dispatch sequence:', $dispatch_sequence ] );
			my $built_ref = $self->_build_elements( $dispatch_sequence, $value_definitions );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Received built ref:", $built_ref ] );
			my	$return .= sprintf(
					$built_ref->{sprintf_string},
					$built_ref->{integer}->{value},
					$built_ref->{decimal}->{value},
				);
			$return = $built_ref->{sign} . $return if $built_ref->{sign} and $return;
			return $return;
		};
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

sub _build_percent_sub{
	my( $self, $type_filter, $list_ref, $conversion_defs ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_percent_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return decimal values",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and code hash ref:', $conversion_defs	] );
	
	my $sprintf_string;
	my $decimal_count = 0;
	# Process once to determine what to do
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( defined $piece->[0] ){
			if( $piece->[0] eq '%' ){
				$sprintf_string .= '%%';
			}elsif( $piece->[0] eq '.' ){
				$sprintf_string .= '.';
			}else{
				$sprintf_string .= '%s';
				$decimal_count++;
			}
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	$conversion_defs->{no_decimal} = 1 if $decimal_count < 2;
	$conversion_defs->{sprintf_string} = $sprintf_string;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	my $dispatch_sequence = $number_build_dispatch->{percent};
	
	my 	$conversion_sub = sub{
			my $adjusted_input = $_[0];
			if( !defined $adjusted_input or $adjusted_input eq '' ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			my	$value_definitions = clone( $conversion_defs );
				$value_definitions->{initial_value} = $adjusted_input;
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'Building scientific output with:',  $conversion_defs,
			###LogSD		'..and dispatch sequence:', $dispatch_sequence ] );
			my $built_ref = $self->_build_elements( $dispatch_sequence, $value_definitions );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Received built ref:", $built_ref ] );
			my $return;
			if( $decimal_count < 2 ){
				$return .= sprintf(
					$built_ref->{sprintf_string},
					$built_ref->{integer}->{value}
				);
			}else{
				$return .= sprintf(
					$built_ref->{sprintf_string},
					$built_ref->{integer}->{value},
					$built_ref->{decimal}->{value},
				);
			}
			$return = $built_ref->{sign} . $return if $built_ref->{sign} and $return;
			return $return;
		};
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

sub _build_scientific_sub{
	my( $self, $type_filter, $list_ref, $conversion_defs ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_scientific_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return decimal values",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and code hash ref:', $conversion_defs	] );
	
	# Process once to determine what to do
	my ( $sprintf_string, $exponent_sprintf );
	$conversion_defs->{no_decimal} = ( exists $conversion_defs->{decimal} ) ? 0 : 1 ;
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( defined $piece->[0] ){
			if( $piece->[0] =~ /(E)(.)/ ){
				$sprintf_string .= $1;
				$exponent_sprintf = '%';
				$exponent_sprintf .= '+' if $2 eq '+';
				if( exists $conversion_defs->{exponent}->{leading_zeros} ){
					$exponent_sprintf .= '0.' . $conversion_defs->{exponent}->{leading_zeros};
				}
				$exponent_sprintf .= 'd';
			}elsif( $piece->[0] eq '.' ){
				$sprintf_string .= '.';
				$conversion_defs->{no_decimal} = 0;
			}elsif( $exponent_sprintf ){
				$sprintf_string .= $exponent_sprintf;
			}else{
				$sprintf_string .= '%s';
			}
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	$conversion_defs->{sprintf_string} = $sprintf_string;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	my $dispatch_sequence = $number_build_dispatch->{scientific};
	
	my 	$conversion_sub = sub{
			my $adjusted_input = $_[0];
			if( !defined $adjusted_input or $adjusted_input eq '' ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			my	$value_definitions = clone( $conversion_defs );
				$value_definitions->{initial_value} = $adjusted_input;
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'Building scientific output with:',  $conversion_defs,
			###LogSD		'..and dispatch sequence:', $dispatch_sequence ] );
			my $built_ref = $self->_build_elements( $dispatch_sequence, $value_definitions );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Received built ref:", $built_ref ] );
			my $return;
			if( $built_ref->{no_decimal} ){
				$return .= sprintf(
					$built_ref->{sprintf_string},
					$built_ref->{integer}->{value},
					$built_ref->{exponent}->{value}
				);
			}else{
				$return .= sprintf(
					$built_ref->{sprintf_string},
					$built_ref->{integer}->{value},
					$built_ref->{decimal}->{value},
					$built_ref->{exponent}->{value}
				);
			}
			$return = $built_ref->{sign} . $return if $built_ref->{sign} and $return;
			return $return;
		};
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

sub _build_fraction_sub{
	my( $self, $type_filter, $list_ref, $conversion_defs ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_fraction_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return integer and fraction strings",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and code hash ref:', $conversion_defs	] );
	
	# I'm worried about pulling the sprintf parser out of here and I may need to put it back sometime
	
	my $dispatch_sequence = $number_build_dispatch->{fraction};
	my $conversion_sub = sub{
		my $adjusted_input = $_[0];
		if( !defined $adjusted_input or $adjusted_input eq '' ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Return undef for empty strings" ] );
			return undef;
		}
		my	$value_definitions = clone( $conversion_defs );
			$value_definitions->{initial_value} = $adjusted_input;
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		'Building scientific output with:',  $conversion_defs,
		###LogSD		'..and dispatch sequence:', $dispatch_sequence ] );
		my $built_ref = $self->_build_elements( $dispatch_sequence, $value_definitions );
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Received built ref:", $built_ref ] );
		my $return;
		if( $built_ref->{integer}->{value} ){
			$return = $built_ref->{integer}->{value};
			if( $built_ref->{fraction}->{value} ){
				$return .= ' ';
			}
		}
		if( $built_ref->{fraction}->{value} ){
			$return .= $built_ref->{fraction}->{value};
		}
		if( !$return and $built_ref->{initial_value} ){
			$return = 0;
		}
		$return = $built_ref->{sign} . $return if $built_ref->{sign} and $return;
		return $return;
	};
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

sub _build_elements{
	my( $self, $dispatch_ref, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached the dispatcher for number building with:', $value_definitions,
	###LogSD			'..using dispatch list', $dispatch_ref	] );
	for my $method ( @$dispatch_ref ){
		$value_definitions = $self->$method( $value_definitions );
		###LogSD		$phone->talk( level => 'debug', message => [
		###LogSD			'Updated value definitions:', $value_definitions, ] );
	}
	return $value_definitions;
}

sub _convert_negative{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_convert_negative', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _convert_negative with:', $value_definitions,	] );
	
	if( $value_definitions->{negative_type} and $value_definitions->{initial_value} < 0 ){
		$value_definitions->{initial_value} = $value_definitions->{initial_value} * -1;
	}
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'updated value definitions:', $value_definitions,	] );
	return $value_definitions;
}

sub _divide_by_thousands{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_divide_by_thousands', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _convert_to_percent with:', $value_definitions,	] );
	if(	$value_definitions->{initial_value} and
		$value_definitions->{divide_by_thousands} ){
		$value_definitions->{initial_value} =
			$value_definitions->{initial_value}/
				( 1000**$value_definitions->{divide_by_thousands} );
	}
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'updated value definitions:', $value_definitions,	] );
	return $value_definitions;
}

sub _convert_to_percent{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_convert_to_percent', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _convert_to_percent with:', $value_definitions,	] );
	
	$value_definitions->{initial_value} = $value_definitions->{initial_value} * 100;
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'updated value definitions:', $value_definitions,	] );
	return $value_definitions;
}

sub _split_decimal_integer{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_split_decimal_integer', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _split_decimal_integer with:', $value_definitions,	] );
	
	# split integer and decimal
	my @results = sprintf( '%0.30f', $value_definitions->{initial_value} ) =~ /^(-)?(\d+)(\.)?(\d+)?/;
	my	$integer = $2;
	my	$decimal = $4;
		$value_definitions->{sign} = $1;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Results:', @results,
	###LogSD		'integer: ' . $integer, 'decimal: ' . $decimal,
	###LogSD		(($1) ? 'sign: ' . $1 : '' ),  ] );
	
	# handle slightly underreported decimals
	if(	$decimal and
		my @results = $decimal =~ /^(0+)?([1-8][0-9]+)?(9{4}9+)([0-9]*[1-9])?(0+)?$/ ){
		$decimal = undef;
		$decimal = $2 if defined $2;
		$decimal .= $3;
		$decimal .= $4 if defined$4;
		my	$adder	= 1;
			$adder .= 0 x length( $4 ) if $4;
		my $stripped_length = length( $decimal );
			$stripped_length += length( $results[0] ) if defined $results[0];
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		'Results:', @results,
		###LogSD		'Stripped decimal: ' . $decimal,
		###LogSD		"Adjusting a potentially underreported decimal with: $adder",  ] );
		$decimal += $adder;
		$decimal = $results[0] . $decimal if defined $results[0];
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		'Resulting decimal: ' . $decimal  ] );
		if( length( $decimal ) > $stripped_length ){
			$integer++;
			$decimal = substr( $decimal, 1, $stripped_length );
		}
	}
	$value_definitions->{integer}->{value} = $integer;
	$value_definitions->{decimal}->{value} = $decimal;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Updated ref:', $value_definitions		] );
	return $value_definitions;
}

sub _move_decimal_point{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_move_decimal_point', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _move_decimal_point with:', $value_definitions,	] );
	my	$exponent;
	if(	$value_definitions->{integer}->{value} and 
		$value_definitions->{integer}->{value} =~ /([1-9])/ ){
		my $stopped = $+[0];
		###LogSD	$phone->talk( level => 'debug', message =>[ "Matched integer value at: $stopped",	] );
		$exponent = length( $value_definitions->{integer}->{value} ) - $stopped;
	}elsif( $value_definitions->{decimal}->{value} and 
			 $value_definitions->{decimal}->{value} =~ /([1-9])/ ){
		my $stopped = $+[0];
		###LogSD	$phone->talk( level => 'debug', message =>[ "Matched decimal value at: -$stopped",	] );
		$exponent =  '-' . $stopped;
	}else{
		$exponent = 0;
	}
	###LogSD	$phone->talk( level => 'debug', message =>[ "Initial exponent: $exponent",	] );
	my	$exponent_remainder = $exponent % $value_definitions->{integer}->{minimum_length};
	###LogSD	$phone->talk( level => 'debug', message =>[ "Exponent remainder: $exponent_remainder",	] );
		$exponent -= $exponent_remainder;
	###LogSD	$phone->talk( level => 'debug', message =>[ "New exponent: $exponent",	] );
		$value_definitions->{exponent}->{value} = $exponent;
	if( $exponent < 0 ){
		my $adjustment = $exponent * -1;
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"The exponent |$exponent| is less than zero - the decimal must move to the right"  ] );
		$value_definitions->{decimal}->{value} =~ /(.{$adjustment})(.+)/;
		$value_definitions->{integer}->{value} = ($1 * 1);
		$value_definitions->{decimal}->{value} = $2;
	}elsif( $exponent > 0 ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"The exponent -$exponent- is greater than zero - the decimal must move to the left"  ] );
		$value_definitions->{integer}->{value} =~ /(.*)(.{$exponent})$/;
		$value_definitions->{decimal}->{value} = $2 . $value_definitions->{decimal}->{value};
		$value_definitions->{integer}->{value} = $1;
	}
	
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Updated ref:', $value_definitions		] );
	return $value_definitions;
}

sub _round_decimal{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_round_decimal', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _round_decimal with:', $value_definitions,	] );
	if( $value_definitions->{no_decimal} ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"No decimal condition identified - rounding integer as needed"  ] );
		if( substr( $value_definitions->{decimal}->{value}, 0, 1 ) > 4 ){
			$value_definitions->{integer}->{value}++;
		}
		delete $value_definitions->{decimal};
	}elsif( $value_definitions->{decimal}->{max_length} ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Enforcing decimal max length: " . $value_definitions->{decimal}->{max_length}  ] );
		if( !$value_definitions->{decimal}->{value} ){
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"Decimal found to be 0 or no string"] );
			$value_definitions->{decimal}->{value} = 
				0 x $value_definitions->{decimal}->{max_length};
		}else{
			$value_definitions->{decimal}->{value} =~ 
				/^(.{$value_definitions->{decimal}->{max_length}})(.)/;
			$value_definitions->{decimal}->{value} = $1;
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"New decimal: " . $value_definitions->{decimal}->{value}] );
			if( $2 and $2 > 4 ){
				###LogSD	$phone->talk( level => 'info', message => [
				###LogSD		"Need to round the decimal up for: $2" ] );
				$value_definitions->{decimal}->{value}++;
				if( length( $value_definitions->{decimal}->{value} ) > 
					$value_definitions->{decimal}->{max_length} 			){
					###LogSD	$phone->talk( level => 'info', message => [
					###LogSD		"Need to round the integer up because: " . $value_definitions->{decimal}->{value} ] );
					$value_definitions->{integer}->{value}++;
					$value_definitions->{decimal}->{value} = 
						substr( $value_definitions->{decimal}->{value}, 1 );
					###LogSD	$phone->talk( level => 'info', message => [
					###LogSD		"New decimal: " . $value_definitions->{decimal}->{value}] );
				}
			}
		}
	}
	
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Updated ref:', $value_definitions		] );
	return $value_definitions;
}

sub _add_commas{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_add_commas', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _add_commas with:', $value_definitions,	] );
	if( exists $value_definitions->{integer}->{comma} ){
		$value_definitions->{integer}->{value} = $self->_add_integer_separator(
			$value_definitions->{integer}->{value},
			$value_definitions->{integer}->{comma},
			$value_definitions->{integer}->{group_length},
		);
	}
	
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Updated ref:', $value_definitions		] );
	return $value_definitions;
}

sub _pad_exponent{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_pad_exponent', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _pad_exponent with:', $value_definitions,	] );
	if(	$value_definitions->{exponent}->{leading_zeros} and 
		length( $value_definitions->{exponent}->{value} ) <
		$value_definitions->{exponent}->{leading_zeros}		 ){
		$value_definitions->{exponent}->{value} =
			0 x ( $value_definitions->{exponent}->{leading_zeros} - 
					length( $value_definitions->{exponent}->{value} ) ) .
			$value_definitions->{exponent}->{value};
	}
	
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Updated ref:', $value_definitions		] );
	return $value_definitions;
}

sub _build_fraction{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_build_fraction', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _build_fraction with:', $value_definitions,	] );
	if( exists $value_definitions->{decimal}->{value} ){
		$value_definitions->{fraction}->{value} = 
			( $value_definitions->{fraction}->{divisor} ) ?
				$self->_build_divisor_fraction( $value_definitions ) :
				$self->_continuous_fraction(
					'0.' . $value_definitions->{decimal}->{value}, 20,
					$value_definitions->{fraction}->{target_length},
				);
	}
	$value_definitions->{fraction}->{value} //= 0;
	if( $value_definitions->{fraction}->{value} eq '1' ){
		$value_definitions->{integer}->{value}++;
		$value_definitions->{fraction}->{value} = 0;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Updated ref:', $value_definitions		] );
	return $value_definitions;
}

sub _build_divisor_fraction{
	my( $self, $value_definitions, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_elements::_build_divisor_fraction', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Reached _build_divisor_fraction with:', $value_definitions,	] );
	my $divisor = $value_definitions->{fraction}->{divisor};
	my $decimal = '0.' . $value_definitions->{decimal}->{value};
	my $low_numerator = int( $divisor * $decimal );
	my $high_numerator = $low_numerator + 1;
	my $low_delta = $decimal - ($low_numerator / $divisor);
	my $high_delta = ($high_numerator / $divisor) - $decimal;
	my $return;
	my $add_denominator = 0;
	if( $low_delta < $high_delta ){
		$return = $low_numerator;
		$add_denominator = 1 if $return;
	}else{
		$return = $high_numerator;
		if( $high_numerator == $divisor ){
			$return = 1;
		}else{
			$add_denominator = 1;
		}
	}
	$return .= "/$divisor" if $add_denominator;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final fraction: $return"		] );
	return $return;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings - Parser of XLSX format strings

=head1 DESCRIPTION

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