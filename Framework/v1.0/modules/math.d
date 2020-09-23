Procedure Area(%w,%l);
     return @Mult(%w,%l);
end Area;

function Log10(%n);
    return @div(@log(%n),@log(10));
end Log10;

function IsFloat(%float);
     return @inttobool(@chrpos(%float,"."));
end IsFloat;

function isOddEven(%number);
   return @eval(%number mod 2);
end isOddEven;

function BitAnd(%a,%b);
   return @eval(%a and %b);
End BitAnd;

function BitOr(%a,%b);
   return @eval(%a or %b);
End BitOr;

function BitNot(%a);
  return @eval(not %a);
end BitNot;

function BitXor(%a,%b);
   return @eval(%a Xor %b);
End BitXor;

function isZero(%aZero);
  return @eval(%aZero = 0);
end isZero;

function Max(%num1,%num2);
var %icheck;

    %iCheck = @eval(%num1 > %num2);
    if (%iCheck) then
        return %num1;
        break;
    else
       return %num2
    end if
end Max;

function Min(%num1,%num2);
var %icheck;

    %iCheck = @eval(%num1 < %num2);
    if (%iCheck) then
        return %num1;
        break;
    else
       return %num2
    end if
end Min;

