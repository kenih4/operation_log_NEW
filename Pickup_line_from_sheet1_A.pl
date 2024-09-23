use strict;
use warnings;
#use utf8;
#use open ":utf8";
#binmode STDIN, ':encoding(cp932)';
#binmode STDOUT, ':encoding(cp932)';
#binmode STDERR, ':encoding(cp932)';

# sheet1.xmlのC列には、sharedStrings.xmlのどの行を当たればよいかが記されているで、その値（何行目か）を基にエクセル上でのC列を再現する
#
#   perl Pickup_line_from_sheet1_A.pl sheet1_columnA.out > sheet1_A_DATEpickuped.out

open(SHEET1FILE,  $ARGV[0]) or die("Error:$!");

my $dt=0;
while(my $line = <SHEET1FILE>){
  chomp($line);
#  print "$line\n";

  if ($line =~ /<(c r="A)(\d+)(".*>)/){   
      my $gyo = $2;
#      print "A$gyo\t";
#対象文字列: 
#   <c r="A1" s="49"><v>45444</v></c>
#   <c r="A2" s="30"/><c r="B2" s="30"/><c r="C2" s="32" t="s"><v>197</v></c>
    if ($line =~ /<(c r="A)(\d+)(" s="\d+">)/){
      $line =~ s/<.*?>//g;
      $dt = $line;
    }
    print "A$gyo\t$dt\n";
  }


}

close(SHEET1FILE);
