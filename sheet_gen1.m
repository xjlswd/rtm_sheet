

function sheet_gen(xls, sheet)
    [~,~,raw] = xlsread(xls,sheet);
    M = raw(5:end-1,:);
    [m,n]= size(M);

    vector = randperm(m);
    new = cell(m,n);

    for i = 1:m
        new(i,:) = M(vector(i),:);
    end
%     raw(5:end-1,:) = new;
    i=1;
    num2 = 0 ;
    num3 = 0;

    while 1 
        if new{i , 3} ==1 
            a =1;
        elseif new{i , 3} ==2
            a = 2;
        elseif new{i , 3}==3
            a = 2;
        elseif new{i , 3} ==4
            a = 2;    
        elseif new{i , 3} >=5 & new{i , 3} <=6
            a = 3;
        else
            a = 4;
        end
        if new{i+1 , 3} ==1 
            b =1;
        elseif new{i+1 , 3} ==2
            b = 2;
        elseif new{i+1 , 3}==3
            b = 2;
        elseif new{i+1 , 3} ==4
            b = 2;    
        elseif new{i+1 , 3} >=5 & new{i , 3} <=6
            b = 3;
        else
            b = 4;
        end
        if new{i+2 , 3} ==1 
            c =1;
        elseif new{i+2 , 3} ==2
            c = 2;
        elseif new{i+2 , 3}==3
            c = 2;
        elseif new{i+2 , 3} ==4
            c = 2;    
        elseif new{i+2 , 3} >=5 & new{i , 3} <=6
            c = 3;
        else
            c = 4;
        end
%         n= ismember([a,b,c],[2,5,6]);
%         p=0;
%         for j = 1:3
%             p = n(j)+p;
%         end
        flag1 = a ~= b & a~=c & b~=c; % & p~=3;
        flag2 = (new{i , 4} ~= new{i+1 , 4}) & (new{i , 4} ~= new{i+2 , 4}) & (new{i +1 , 4} ~= new{i+2 , 4});
         
        flag3 = (new{i,2} == '有')& (new{i+1,2} == '有')& (new{i+2,2} == '有');

        if flag1 & flag2  & flag3
            if new{i , 4} == 'B'
                if new{i+1 ,4} == 'C' | new{i+2 ,4} == 'C'
                    i = i+1;
                    continue;
                end
            end
            if new{i , 4} == 'C'
                if new{i+1 ,4} == 'B' | new{i+2 ,4} == 'B'
                    i = i+1;
                    continue;
                end
            end
            if new{i , 4} == 'E'
                if new{i+1 ,4} == 'F' | new{i+2 ,4} == 'F'
                    i = i+1;
                    continue;
                end
            end
            if new{i , 4} == 'F'
                if new{i+1 ,4} == 'E' | new{i+1 ,4} == 'G' |new{i+2 ,4} == 'E' | new{i+2 ,4} == 'G' 
                    i = i+1;
                    continue;
                end
            end
            if new{i , 4} == 'G'
                if new{i+1 ,4} == 'F'|new{i+1 ,4} == 'F'
                    i = i+1;
                    continue;
                end
            end
                
            new{i , 4} = ['*',new{i,4}];
            new{i+1 , 4} =['*',new{i+1,4}];
            new{i+2 , 4} =['*',new{i+2,4}];
            i = i +4;
            num3 = num3 +1 ;
            if num3 == 2
                break;
            end
        end
        i = i+1;
        if i ==m-3
            break;
        end
    end
    num3
    while 1 
        if new{i , 3} ==1 
            a =1;
        elseif new{i , 3} >=2 & new{i , 3} <=4
            a = 2;
        elseif new{i , 3} >=5 & new{i , 3} <=6
            a = 3;
        else
            a = 4;
        end
        if new{i+1 , 3} ==1 
            b =1;
        elseif new{i+1 , 3} >=2 & new{i , 3} <=4
            b = 2;
        elseif new{i+1 , 3} >=5 & new{i , 3} <=6
            b = 3;
        else
            b = 4;
        end
        
        
        flag1 = (a ~= b);
        flag2 = (new{i , 4} ~= new{i+1 , 4});
         
        flag3 = (new{i,2} == '有')& (new{i+1,2} == '有');

        if flag1 & flag2  & flag3
            if new{i , 4} == 'B'
                if new{i+1 ,4} == 'C'
                    i = i+1;
                    continue;
                end
            end
            if new{i , 4} == 'C'
                if new{i+1 ,4} == 'B'
                    i = i+1;
                    continue;
                end
            end
            if new{i , 4} == 'E'
                if new{i+1 ,4} == 'F'
                    i = i+1;
                    continue;
                end
            end
            if new{i , 4} == 'F'
                if new{i+1 ,4} == 'E' | new{i+1 ,4} == 'G'
                    i = i+1;
                    continue;
                end
            end
            if new{i , 4} == 'G'
                if new{i+1 ,4} == 'F'
                    i = i+1;
                    continue;
                end
            end
                
            new{i , 4} = ['*',new{i,4}];
            new{i+1 , 4} =['*',new{i+1,4}];
            i = i +3;
            num2 = num2 +1 ;
            if num2 == 2
                break;
            end
        end
        i = i+1;
        if i ==m-3
            break;
        end
    end
    num2
 
    new(:,1) = M(:,1);
    excelData = cell(m+1,8);
    excelData(1,:)={'序号','测试类型','违禁物品或其模拟物代号','身体位置代号'...
        ,'自动探测检出','自动探测误报','人工判图检出','人工判图误报'};
    excelData(2:end,:) = new;
    xlswrite([pwd,'\new\',xls],excelData,sheet);
    
end


