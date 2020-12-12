use utf8;


binmode STDOUT, ":utf8";
binmode STDIN, ":utf8";

use String::Approx 'adist';

use Encode qw(decode);
use Spreadsheet::WriteExcel;
use Spreadsheet::ParseExcel::FmtUnicode;

use Data::Dumper;
use Getopt::Long;
use Spreadsheet::Read qw(ReadData);
use Unicode::Normalize;

my %present;

sub make_clean{
	my $str = shift @_;
	$str = uc($str);
	$str =~ s/^\s+//;
	$str =~ s/\s+$//;
	$str =~s/ +/ /;
	$str =~ s/Á/A/;
	$str =~ s/É/E/;
	$str =~ s/Í/I/;
	$str =~ s/Ó/O/;
	$str =~ s/Ú/U/;
	$str =~ s/Ñ/N/;
	#$str =~ s/[^A-ZÑ ,]//g;
	$str =~ s/^\s+//;
	$str =~ s/\s+$//;

	return $str;
}

sub clean_extracted_name {
	my $str = shift @_;
	$str =~ s/[:)(]//;

	return $str;
}

sub check_if_in_name {
	my $test = shift @_;

	my @elements = split / /, $test;
	my $n_words_in_test = @elements;

	my $true_name_present = '';

	my $maximum_matches = 2;

	my $max_match = 0;
	my $max_name = "";

	for $name (@names){

		my $count = 0;
		@words_in_name = split / /, $name;

		for my $element (@elements){
			
			if($name =~ qr/$element/ ){
				$count++;
				
			}
		}

		if( $count > $max_match ) {
			$max_match = $count;
			$max_name = $name;
		} 

	}

	if( $max_match >= 2 ){
		$present{$max_name} ++;
	} else {
		$max_name = '';
	}


	return $max_name;
}

sub approximate_length{
	my $a = shift;
	my $b = shift;
	my $len_a = length $a;
	my $len_b = length $b;
	my $max_difference = 1;

	if( abs( $len_a - $len_b ) <= $max_difference ){
		return 1;
	} else {
		return 0;
	}

}

sub check_if_in_name_fuzzy {
	my $test = shift @_;

	my @elements = split / /, $test;
	my $n_words_in_test = @elements;
	my $flag_match = 0;

	print "---> n_words_in_test : $n_words_in_test \n";

	my $true_name_present = '';
	my $maximum_matches = 2;

	for my $name (@names){

		my $count = 0;
		my @words_in_name = split / /, $name;

		

		print "DOING FUZZY SEARCH for $test AT: $name\n";
		for my $element (@elements){
			
			for my $word (@words_in_name){
				print "		Comparing $word vs $element ";
				if(     ( abs( adist( $element, $word) )  <= 1 ) &&  ( approximate_length($element, $word) == 1 )  ) {
					$count++;
					print " --- MATCH! \n";
				}else {
					print "\n";
				}
			}

			print "---> Count was: $count\n";

			if( $count >= $maximum_matches ) {
				$present{$name} ++;
				$flag_match ++;
				$true_name_present = $name;
			} 
		}				
	}
	return $true_name_present;
}


my $file_name = "";
my $selected_group = "2C";

GetOptions( "g=s" => \$selected_group, 
			"o=s" => \$file_name,
			"d=s" => \$selected_date,
			"zs" => \$zoom_script_mode,
      "s=s" =>\$search_student,
      "f=s" => \$file_to_read_students
			) 
		or die ("Error in command line");


print("You want to open $file_name\n");
my $book = ReadData ( $file_name );


#print Dumper $book;
#$jjjj = <>;
print "The number of sheets is : " . $book->[0]{sheets}. "\n";

$number_of_sheets = $book->[0]{sheets};

$names_of_sheets = $book->[0]{sheet};

if( $search_student ) {

  print " You are looking for $search_student \n ";

  
  $index = 0;
  my @arr_index = (0 .. $number_of_sheets);

  my @results;

  foreach $group(keys %$names_of_sheets){
  ##//$arr_index[$$names_of_sheets{$group}] = ;
    $sheet = $book->[$$names_of_sheets{$group}];
   
    ## print Dumper $sheet;

    $maxrow = $$sheet{maxrow};
    $maxcol = $$sheet{maxcol};

    ##print "This sheet has this many rows:" . $$sheet{maxrow} . "\n";
    ##print "This sheet has this many Columns:" . $$sheet{maxcol} . "\n";

    for $i  (1 .. $maxrow){
      my %person_res;
	    my $name =  $$sheet{"B".$i};
      if( $name =~ /\Q$search_student/gi ){
         print "In group: $group \n";
          print "Found person: $name\n"; 
         $person_res{name} = $name;
         $person_res{group_index} = $$names_of_sheets{$group};
         $person_res{group} = $group;
         $person_res{row} = $i;
         # get the columns, the absenses 
         my @absenses;
  
    print "This sheet has this many Columns:" . $$sheet{maxcol} . "\n";
    #print Dumper $sheet;
    #     $valueee = <>;
         my @dates;
         for $j ( 2 .. $maxcol ){
           push( @absenses, $sheet->{cell}[$j][$i] );
           push( @dates, $sheet->{cell}[$j][1]); 
           print  $sheet->{cell}[$j][1].  " --  "; 
           print  $sheet->{cell}[$j][$i].  " --  \n"; 
         }
         $person_res{absenses} = @absenses;
         $person_res{dates} = @dates;
          push(@results, %person_res);
      }
    }

  }

  ## print Dumper @results;
  print "These are the results:\n";
  $count = 1;
  for $res (@results) {
    print $count . ". ";
    @info = $res->{absenses}


  }

  exit;
}







print("Selected group: $selected_group\n");
$index = 0;
my @arr_index = (0 .. $number_of_sheets);

foreach $name (keys %$names_of_sheets){
	print "Got: $name \n";
	if( $name eq $selected_group){
		$index  = $$names_of_sheets{$name};
	}
	$arr_index[$$names_of_sheets{$name}] = $name;
	
}

print "\n\n Index is: $index \n";

print Dumper $book->[$index];

$sheet = $book->[$index];

$maxrow = $$sheet{maxrow};
$maxcol = $$sheet{maxcol};

print "This sheet has this many rows:" . $$sheet{maxrow} . "\n";
print "This sheet has this many Columns:" . $$sheet{maxcol} . "\n";

## if you are looking for a student


print "These are all the names of students:\n";

for $r (2 .. $maxrow){
	my $name =  $$sheet{"B".$r};
	print $name . "\n";
	push @names, make_clean($name);
}




print "Enter data:";
my %present;
my %unmatched;
my $name_in_title = '';
my $line_count = 0;
my %total_names_in_titles; # contains the name of the students


if( $file_to_read_students ) {
open(FH,"<", $file_to_read_students ) or die "I couldn't open $file_to_read_students\n"; 
$file_handle = FH;
}else {
$file_handle = STDIN;
}

while($in = <$file_handle>){
	# if( $in =~ m/@names/ ){
	# 	print "here!";
	# }

	chomp($in);

	if( $in eq "" ){
		last;
	}

	print "\n_________________Doing line: " . ++$line_count . "___________________\n";
	

	if( $in =~ /^From (.+) to/ ){
		## chat title
		print "\n+Chat log title: $1 \n";
		$name_in_title = $1;
		$name_in_title = make_clean($name_in_title);

	}else {
	## got to parse From González Samantha to Me: (Privately) (8:49 AM)
		
		$in = make_clean($in);

		if($zoom_script_mode == 1){ ## if -zs is placed on the command line argument
			#input is: 
			# 11:08:23	 From zavala ivana to Valderrama Jesús (Privately) : ZAVALA IVANA, PRESENT
			# from log file
			if( $in =~ /\sto\s.+\s:/ ){
				print "FROM HOST!\n ";
				$in =~ /From (.+) to .*: (.+)$/i;
				$name_in_title = make_clean($1);
				$test = make_clean($2);
				$test =~ s/\W+(PRES.+|PES.+|PRS.+)$//;
			} else {
				print "no HOST \n ";
				$in =~ /From (.+) : (.+)$/i;
				$name_in_title = make_clean($1);
				$test = make_clean($2);
				$test =~ s/\W+(PRES.+|PES.+|PRS.+)$//;
			}

			print "-> name_in_title: $name_in_title \n";
			print "-> test: $test \n";
 


		} else {  ## directly from chat text
			

			$res = ( $in =~ /^(\w+\s+\w+)\W+(PRESENTE|PRESENT|PRESENE|PRESEN|PRE.*)$/i );

			$test = make_clean( $in );

			#$test =~ s/\W+(PRES.*|PES.*|PRS.*)$//;
			$test =~ s/\W+\w+$//;

			# if( $res == '' && $name_in_title != '') {  ## if I can't get this, use the title
			# 	$test = uc( $name_in_title );
			# 	$test  = make_clean($test);
			# }else {
			# 	$test = uc($1);
			# 	$test = make_clean($test); 
			# }



		}

		## if they placed here... 

		$test =~ s/\W+HERE$//;


		$test = clean_extracted_name($test);
		$name_in_title = clean_extracted_name($name_in_title);
		
		print ">> $in\n";
		print "Matched: $test \n";
		print "res : $res \n";

		$test =~s/ +/ /;
		my $flag_match = 0;
		
		####
		print "--TESTING FOR: $test \n";

		my $result = check_if_in_name( $test );

		print ">> RESULT IS: $result \n";

		$total_names_in_titles{$name_in_title} = 1;

		if( $result eq ""){  ## if there were no matches with regex, then make an approximate search

			print ("---FUZZY SEARCH WITH TEST FOR $test \n");
			$result = check_if_in_name_fuzzy( $test );


			####

			if( $result eq ""){
				print "----DIRECT SEARCH WITH name_in_title : $name_in_title \n";
				$result = check_if_in_name( $name_in_title );

				if( $result eq ""){ 
					print "----- FUZZY SEARCH WITH name_in_title : $name_in_title\n";
					$result = check_if_in_name_fuzzy( $name_in_title );

					if( $result eq ""){ 
						print "------ X OOPS... this was the LAST chance... this did not match at all.. \n";
						$unmatched{$test}++;
					} else {
						print "Going to save $result \n";
						$present{$result}++;
					}

				} else {

					
					print "Going to save $result \n";
					$present{$result}++;
				}
				
			} else {
				print "Going to save $result \n";
				$present{$result}++;
			}		

		} else {
			print "Going to save $result \n";
			$present{$result}++;
		}
		
	}
	print "\n";

}

## just for formating 
$max = 0;
for (@names){
	$len = length $_;
	if($len > $max) {
		$max = $len;
	}
}

my @assistance_array;

sub show_list {
	$list_n = 1;
	$total_present = 0;
	$total_absent = 0;
	my $i = 0;
	for (@names){
		if( $list_n < 10){
			print " $list_n. ";
		}else {
			print "$list_n. ";
		}
		print "$_  ";
		$len = length $_;
		$spaces = "." x ( $max - $len ) ;
		print $spaces;
		if(exists $present{$_} ){
			print " PRESENT \n";
			$total_present++;
			$assistance_array[$i] = "0";
		} else {
			print "    x    \n";
			$total_absent++;
			$assistance_array[$i] = "1";
		}
		$list_n++;
		$i++;
	}
	print "  # Present: $total_present   # Absent: $total_absent \n\n";

} 

# calculate the total lines that were set
my $n_lines = keys %total_names_in_titles;
print "TOTAL AMOUNT OF LINES ENTERED: $n_lines\n";

show_list();
if ($n_lines != $total_present){
	print "PLEASE CHECK ... \n" ;
} else {
	print "CHECKSUM CORRECT \n";
}



if( keys %unmatched > 0 ){

	print "________________________________ CHECK THIS OUT: _____________________ \n";
	for (sort keys %unmatched){
		print "$_  : " . $unmatched{$_} . "\n";
		print "[MANUAL] Give me the list number of this student: ";
		my $list_n = <>;
		chomp($list_n);
		$present{ $names[$list_n - 1 ] }++;
		$assistance_array[$list_n] = "0";
	} 
	show_list();
	if ($n_lines != $total_present){
		print "PLEASE CHECK ... \n" ;
	} else {
		print "CHECKSUM CORRECT \n";
	}

} 


print "   Do you want to modify this? [m] or ANY KEY to continue ";

$input = <STDIN>;
chomp($input);

if( $input =~/m/ ){
 while(1){
   print "List number to toggle: [number to continue] [any character to finish]\n";
    my $list_n = <>;
    chomp($list_n);

    if( $list_n =~/\d+/ ){ 
      my $current_value =  $assistance_array[$list_n]  ;
      print "This has: $current_value \n";
      if(  !exists(  $present{ $names[$list_n - 1 ]  } ) ){

      # $assistance_array[$list_n] = "1";

		    $present{ $names[$list_n - 1 ] }++;
      } else {

      #$assistance_array[$list_n] = "1";

        delete( $present{ $names[$list_n - 1 ] });
      }
        
      show_list();


    } else {
       ## want to exit
       last;
       
    }
  }

}

print "   Do you want to save this? [y] ";
$input = <STDIN>;
chomp($input);

($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime();
$year+=1900;
$mon++;

if( $input =~ /y/ or $input eq ""){
	print "Saving... \n";

	my $workbook = Spreadsheet::WriteExcel->new($file_name);
	my $fmt = $workbook->add_format();


	for $i_sheet ( 1 .. $number_of_sheets ){ #go through sheets
		$sheet = $book->[$i_sheet];

		$maxrow = $$sheet{maxrow};
		$maxcol = $$sheet{maxcol};

		print "($maxrow , $maxcol)\n";

		$worksheet = $workbook->add_worksheet( $arr_index[$i_sheet] ); # create sheet
		$worksheet->set_column('A:A', 3);
		$worksheet->set_column('B:B', 40);
		$worksheet->set_column('C:Z', 8);

		for $row (1 .. $maxrow){
			for $col (1 .. $maxcol){
				$worksheet->write($row-1, $col-1, $book->[$i_sheet]{cell}[$col][$row]);
			
			} 



			# then write an extra column for the new info
			# place the date on the first column
			if( $index == $i_sheet && $row == 1){
				if( $selected_date eq ''){
					$worksheet->write($row-1, $maxcol, "$mday/$mon/$year" );
				} else {
					$worksheet->write($row-1, $maxcol, $selected_date );
				}
			}
			if( $index == $i_sheet && $row > 1){
				if( $assistance_array[$row-2 ] == "0"){ 
					$worksheet->write($row-1, $maxcol, $assistance_array[$row-2 ] );
				}else {
					$worksheet->write($row-1, $maxcol, $assistance_array[$row-2 ]);
				}
			}
		}
	}

	
	
	
	$workbook->close();

	
}



print "Zome mode is ON \n" if $zoom_script_mode;
print "DONE!\n\n";
