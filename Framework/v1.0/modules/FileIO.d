function openfile($lzfile);
var $buff;

    open [$lzfile] for binary;
           $buff = @space(@lof());
           Get,,$buff;
    closefile;
    
    return $buff;
end openfile;

function writefile($lzfile,$buff);
var $buff;

    open [$lzfile] for binary;
           write,,$buff;
    closefile;

end writefile;

function FixPath($lzPath);
  if (@right($lzPath,1) == "\") then
     return $lzPath;
  else
    return $lzPath & "\";
 end if

end FixPath;

function GetWindowsDirectory();
var %a,$b;
    $b = @space(215);
    %a = @Win32("GetWindowsDir",$b,215);
    $b = @left($b,%a);
    return $b;
    $b = "";
end GetWindowsDirectory;

