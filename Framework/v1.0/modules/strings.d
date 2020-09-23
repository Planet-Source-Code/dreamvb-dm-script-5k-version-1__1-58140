function CountIF($s,$t);
var %num,%i;
var $ch;
  %num = 0;
  for %i = 1 to @len($s);
     $ch = @StrCpy($s,%i,1);
     if ($ch == $t) then
         %num = @add(%num,1)
      end if
  next
  $ch = "";
  %i = 0;
  return %num;
  %num = 0;
  $s = "";
end CountIf;

function StrRev($dwStr);
var $dwStr,$s,%v;

%v = 1

for %v = 1 to @len($dwStr);
  $s = "" & @strcpy($dwStr,%v,1) & $s;
next

return $s;
$s = "";

end StrRev;

function Split($Str,$a,%Index);
var $Str,$ch,$b;
var %i,%Cnt;
var $List();

%Cnt = 0;

for %i = 1 to @len($Str);
   $ch = @StrCpy($Str,%i,1);
   $b = "" & $b & $ch;

   if ($ch == $a) then
     %cnt = @Add(%cnt,1);
     Redim $List(%cnt);
     $b = @left($b,@Sub(@len($b),1));
     $List(%cnt) = $b;
     $b = "";
   end if
next

$ch = "";
$b = "";
%cnt = 0;

return $List(%Index);

Destroy $List;

End Split;

