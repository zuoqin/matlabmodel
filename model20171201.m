lastwarn('');
filename = 'E:/DEV/matlab/sberpb/Modelling20171205.xlsx';
sheet = 1;
xlRange = 'H3:I377';

subsetA = xlsread(filename,sheet,xlRange);

xlRange = 'AG3:AG377';
colAG = xlsread(filename,sheet,xlRange);

xlRange = 'A3:A377';
dates = xlsread(filename,sheet,xlRange);


[m,n] = size(subsetA);

logs = zeros(m,1);

for k = 3:m
  logs(k) =  log(subsetA(k,2)/subsetA(k-1,2));
end


colK = zeros(m,1);
for k = 52:m
  if k == 52
    start = 2;
  else
    start = k - 51;
  end
  colK(k) =  std(logs (start:k))*sqrt(52);
end

colL = zeros(m,1);
for k = 52:m
  colL(k) =  subsetA(k,1)/subsetA(k,2) - 1.0;
end

colM = zeros(m,1);
for k = 52:m
  colM(k) =  colL(k)/colK(k);
end


colN = zeros(m,1);
colO = zeros(m,1);
for k = 64:m
  colN(k) =  mean(colM (k-12:k));
  colO(k) =  std(colM (k-12:k));
end


p1 = 2;
p2 = 4;
p3 = 0.3;
p4 = 2;
p5 = 3;

incase = 0;
casesmdl = zeros(175560,7);  %11*7*19*6*20 = 175560
% datetime('now')


for p1 = 1:0.2:3            %11 cases
  for p2 = 3.5:0.25:5       %7 cases
    for p3 = 0.1:0.05:1     %19 cases
      for p4 = 1:1:6        %6 cases
        for p5 = 0.2:0.2:4  %20 cases
          found = 0;
          incase = incase + 1;

          colP = zeros(m,1);
          colQ = zeros(m,1);
          for k = 63:m
            colP(k) =  colN(k) + p1 * colO(k);
            colQ(k) =  colN(k) - p1 * colO(k);
          end

          colY = zeros(m,1);
          for k = 68:m
            colY(k) = -colL(k)/p5;
          end

          colR  = zeros(m,1);
          colS  = zeros(m,1);
          colT  = zeros(m,1);
          colU  = zeros(m,1);
          colV  = zeros(m,1);
          colZ  = zeros(m,1);
          colAA = zeros(m,1);

          colAB = zeros(m,1);
          colAD = zeros(m,1);
          colAE = zeros(m,1);
          colAF = zeros(m,1);

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

            if colV(k) ~= 0
            	found = 1;
            end




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

          if found == 1
            mdl = fitlm(colV(80:m),dates(80:m));
            if length(lastwarn()) > 0
              return
            else
              casesmdl(incase,1) = mdl.Rsquared.Ordinary;
              casesmdl(incase,2) = colV(m);

              casesmdl(incase,3) = p1;
              casesmdl(incase,4) = p2;
              casesmdl(incase,5) = p3;
              casesmdl(incase,6) = p4;
              casesmdl(incase,7) = p5;

              %disp(casesmdl(incase,:));
            end
          else
            casesmdl(incase,1) = 0;
            casesmdl(incase,2) = 0;
          end
          

          if mod(incase,1000) == 0
            datetime('now')
            incase = incase
          end

          
        end
      end
    end
  end
end
% datetime('now')


B = sortrows(casesmdl,1);
C = B(170000:175560,1:7);
D = sortrows(C,2);
dlmwrite('e:/data/exp.txt',D);
