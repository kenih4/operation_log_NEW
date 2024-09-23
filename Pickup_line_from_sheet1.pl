use strict;
use warnings;
#use utf8;
#use open ":utf8";
#binmode STDIN, ':encoding(cp932)';
#binmode STDOUT, ':encoding(cp932)';
#binmode STDERR, ':encoding(cp932)';

# sheet1.xmlのC列には、sharedStrings.xmlのどの行を当たればよいかが記されているで、その値（何行目か）を基にエクセル上でのC列を再現する
#
# perl Pickup_line_from_sheet1.pl sheet1_tag_c.out sharedStrings.out  sheet1_A_DATEpickuped.out

open(SHEET1FILE_C,  $ARGV[0]) or die("Error:$!");

open(SHAREDSTRINGSFILE, $ARGV[1]) or die("Error:$!");
my @ss = <SHAREDSTRINGSFILE>;
my $length_of_ss = @ss;
#print "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~length_of_ss = $length_of_ss\n";

#open(SHEET1FILE_A, $ARGV[2]) or die("Error:$!");
#my @sheet1a_date = <SHEET1FILE_A>;
#print @sheet1a_date;
#print $sheet1a_date[5];



while(my $line = <SHEET1FILE_C>){
  chomp($line);
#  print "$line\n";
#対象文字列: <c r="C3199" s="32" t="s"><v>222</v></c>
  if ($line =~ /<(c r="C)(\d+)(".*>)/){        
#    print "対象文字列: $line\nマッチ部分: $2\n\n";
      my $gyo = $2;
      $line =~ s/<.*?>//g;
      #print "$gyo $line\t\t\t";
      print "C$gyo    sharedStrings:$line\t\t\t";

      if ($line < 1) {
        print "\n";#print "=UNDER 1=\n";
      }elsif($line>$length_of_ss){
        print "\n";#print "=OVER LINE $length_of_ss=\n";
      }else{
        print "$ss[$line]";
      }



  }


}


close(SHEET1FILE_C);
close(SHAREDSTRINGSFILE);
#close(SHEET1FILE_A);