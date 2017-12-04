filename = 'e:/data/Modelling20171201.xlsx';
sheet = 1;
xlRange = 'H3:I519';

subsetA = xlsread(filename,sheet,xlRange)

xlRange = 'AG3:AG519';
colAG = xlsread(filename,sheet,xlRange)


[m,n] = size(subsetA)

logs = zeros(m);

for k = 2:m
  logs(k) =  log(subsetA(k,2)/subsetA(k-1,2));
end

subsetA = [subsetA logs]


colK = zeros(m);
for k = 51:m
  arr = logs (k-50:k);
  colK(k) =  std(arr)*sqrt(52);
end

colL = zeros(m);
for k = 51:m
  colL(k) =  subsetA(k,1)/subsetA(k,2) - 1.0;
end

colM = zeros(m);
for k = 51:m
  colM(k) =  colL(k)/colK(k);
end


colN = zeros(m);
colO = zeros(m);
for k = 63:m
  arr = colM (k-13:k);
  colN(k) =  mean(arr);
  colO(k) =  std(arr);
end


p1 = 2;
p2 = 4;
p3 = 0.3;
p4 = 2;
p5 = 3;
colP = zeros(m);
colQ = zeros(m);
for k = 63:m
  colP(k) =  colN(k) + p1 * colO(k);
  colQ(k) =  colN(k) - p1 * colO(k);
end

colY = zeros(m);
for k = 68:m
  colY(k) =  colL(k)/p5;
end

colU = zeros(m);
for k = 67:m
  colU(k) =  colS(k)*( subsetA(k,2) - subsetA(k-1,2) );
end



colR  = zeros(m);
colS  = zeros(m);
colT  = zeros(m);
colV  = zeros(m);
colZ  = zeros(m);
colAA = zeros(m);
for k = 65:m
  if colR(k) + colR(k-1) != 0
    colZ(k) = 0;
  else
    if colV(k-1) - colV(k-2) - colV(k-3) != 0
       if colZ(k-1) != 0
         colZ(k) = 0;
       else
         if colAF(k) < colY(k)
           colZ(k) = -colS(k-1);
         else
           colZ(k) = 0;
         end
       end
    end
  end

  colAA(k) = sum(colZ (67:k-1));
  colS(k) =  sum(colR (63:k-1)) + colAA(k);

  if colS(k) + colR(k-1) > p4
    if colM(k) < p3
      colR(k) = -colS(k-1) - colR(k-1) - colZ(k-1);
    else
       colR(k) = 0.0;
    end
  else
    if colM(k) < p3
      if colS(k-1) = 0
        colR(k) = 0;
      else 
        colR(k) = -colS(k-1) - colR(k-1) - colZ(k-1);
      end
    else
       if colAG(k) > p2
         if colM(k) > colP(k)
            colR(k) = 1;
          else
            colR(k) = 0;
          end
        else
          colR(k) = 0;
       end
    end    
  end

  colT(k) = abs(colR (k)) + 0.002 * subsetA(k,2);
  colV(k) = sum(colU(67:k)) - sum(colT(67:k));

  
end
