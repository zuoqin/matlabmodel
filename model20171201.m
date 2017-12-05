filename = 'E:/DEV/matlab/sberpb/Modelling20171201.xlsx';
sheet = 1;
xlRange = 'H3:I519';

subsetA = xlsread(filename,sheet,xlRange);

xlRange = 'AG3:AG519';
colAG = xlsread(filename,sheet,xlRange);

xlRange = 'A3:A519';
dates = xlsread(filename,sheet,xlRange);


[m,n] = size(subsetA);

logs = zeros(m);

for k = 3:m
  logs(k) =  log(subsetA(k,2)/subsetA(k-1,2));
end


colK = zeros(m);
for k = 52:m
  if k == 52
    start = 2;
  else
    start = k - 51;
  end
  colK(k) =  std(logs (start:k))*sqrt(52);
end

colL = zeros(m);
for k = 52:m
  colL(k) =  subsetA(k,1)/subsetA(k,2) - 1.0;
end

colM = zeros(m);
for k = 52:m
  colM(k) =  colL(k)/colK(k);
end


colN = zeros(m);
colO = zeros(m);
for k = 64:m
  colN(k) =  mean(colM (k-12:k));
  colO(k) =  std(colM (k-12:k));
end


p1 = 2;
p2 = 4;
p3 = 0.3;
p4 = 2;
p5 = 3;

incase = 1;
casesmdl = zeros(184680,1);  %15*6*18*6*19 = 184680

for p1 = 1:0.2:3  %15 cases
  for p2 = 3.5:0.25:5   %6 cases
    for p3 = 0.1:0.05:1  %18 cases
      for p4 = 1:1:6   %6 cases
        for p5 = 0.2:0.2:4  %19 cases



          colP = zeros(m);
          colQ = zeros(m);
          for k = 63:m
            colP(k) =  colN(k) + p1 * colO(k);
            colQ(k) =  colN(k) - p1 * colO(k);
          end

          colY = zeros(m);
          for k = 68:m
            colY(k) = -colL(k)/p5;
          end


          colR  = zeros(m);
          colS  = zeros(m);
          colT  = zeros(m);
          colU  = zeros(m);
          colV  = zeros(m);
          colZ  = zeros(m);
          colAA = zeros(m);

          colAB = zeros(m);
          colAD = zeros(m);
          colAE = zeros(m);
          colAF = zeros(m);

          for k = 65:m




            colAA(k) = sum(colZ (68:k-1));
            colS(k) =  sum(colR (64:k-1)) + colAA(k);

            colU(k) =  colS(k)*( subsetA(k,2) - subsetA(k-1,2) );

            if colS(k) + colR(k-1) > p4
              if colM(k) < p3
                colR(k) = -colS(k-1) - colR(k-1) - colZ(k-1);
              else
                 colR(k) = 0.0;
              end
            else
              if colM(k) < p3
                if colS(k-1) == 0
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

            colT(k) = abs(colR (k)) * 0.002 * subsetA(k,2);
            colV(k) = sum(colU(68:k)) - sum(colT(68:k));




            if colR(k) > 0
              colAD(k) = subsetA(k,2);
            else
              colAD(k) = colAD(k-1);
            end
            
            
            if colS(k) == 0
              if colR(k) == 1
                colAE(k) = subsetA(k,2);
              else
                colAE(k) = 0;
              end
            else
              if (colAD(k) - colAD(k-1) ~= 0) && (colS(k) > 0)
                if colS(k) + colR(k) == 1
                  colAE(k) = colAD(k);
                else
                  if colS(k) + colR(k) > 1
                    colAE(k) = (colAD(k) + colAE(k-1)) / min(colS(k) + colR(k), 2);
                  else
                    if colS(k) + colR(k) > 2
                      colAE(k) = (colAD(k) + colAE(k-1) * colS(k)) / (colS(k) + colR(k));
                    else
                      colAE(k) = 0
                    end
                  end
                end
              else
                colAE(k) = colAE(k-1);
              end
            end
            colAB(k) = colS(k)*(subsetA(k,2) - colAE(k));
            if colAE(k) == 0
              colAF(k) = 0;
            else
              colAF(k) = colAB(k) / colAE(k);
            end

            if colR(k) + colR(k-1) ~= 0
              colZ(k) = 0;
            else
              if colV(k-1) - colV(k-2) - colV(k-3) ~= 0
                 if colZ(k-1) ~= 0
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

          end

          try
            mdl = fitlm(colV(80:517),dates(80:517));
          catch

          

          casesmdl(incase,1) = mdl.Rsquared.Ordinary;
          incase = incase + 1;

          if mod(incase,10000) == 0
            incase = incase
            return
          end
        end
      end
    end
  end
end


