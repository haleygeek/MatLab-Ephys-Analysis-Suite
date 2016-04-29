% -----------------------------------------------------------------------%
% Author: Haley E. Speed, PhD                                            %
% Department of Neurology                                                %
% University of Texas Southwestern Medical Center                        %
% Dallas, TX                                                             %   
%-------------------------------------------------------------------------

%Quick Script to input a column number and transcibe that into a column
%letter for excel

%To use this script define the variable "column_number" before you call
%excelcols.m

% For example, in the .m file calling this one, I would write:
    % column_number = current_col
    % excelcols
    
% The letter value can then be retrieved from the 'column letter' variable
% in the main file. For example:
        % start_xlrange = A1           
        % column_number = max_col
        % excelcols 
        % end_xlrange = strcat(column_letter, num2string(max_row))
        % xlRange = start_xlrange:end_xlrange
        % xlswrite(filename,variable,sheet,xlRange)

        

switch column_number
    case 1
         column_letter = 'A'
    case 2
         column_letter = 'B'
    case 3
        column_letter = 'C'
    case 4
        column_letter = 'D'
    case 5
        column_letter = 'E'
    case 6 
        column_letter = 'F'
    case 7 
        column_letter = 'G'
    case 8
        column_letter = 'H'
    case 9 
        column_letter = 'I'
    case 10
        column_letter = 'J'
    case 11
        column_letter = 'K'
    case 12
        column_letter = 'L'
    case 13
        column_letter = 'M'  
    case 14
        column_letter = 'N'
    case 15 
        column_letter = 'O'
    case 16 
        column_letter = 'P'
    case 17 
        column_letter = 'Q'
    case 18 
        column_letter = 'R'
    case 19
        column_letter = 'S'
    case 20
        column_letter = 'T'
    case 21 
        column_letter = 'U'
    case 22
        column_letter = 'V'
    case 23
        column_letter = 'W'
    case 24
        column_letter = 'X'
    case 25
        column_letter = 'Y'
    case 26
        column_letter = 'Z'
    case 27
         column_letter = 'AA'
    case 28
         column_letter = 'AB'
    case 29
        column_letter = 'AC'
    case 30
        column_letter = 'AD'
    case 31
        column_letter = 'AE'
    case 32 
        column_letter = 'AF'
    case 33 
        column_letter = 'AG'
    case 34
        column_letter = 'AH'
    case 35 
        column_letter = 'AI'
    case 36
        column_letter = 'AJ'
    case 37
        column_letter = 'AK'
    case 38
        column_letter = 'AL'
    case 39
        column_letter = 'AM'  
    case 40
        column_letter = 'AN'
    case 41 
        column_letter = 'AO'
    case 42 
        column_letter = 'AP'
    case 43 
        column_letter = 'AQ'
    case 44 
        column_letter = 'AR'
    case 45
        column_letter = 'AS'
    case 46
        column_letter = 'AT'
    case 47 
        column_letter = 'AU'
    case 48
        column_letter = 'AV'
    case 49
        column_letter = 'AW'
    case 50
        column_letter = 'AX'
    case 51
        column_letter = 'AY'
    case 52
        column_letter = 'AZ'
    case 53
         column_letter = 'BA'
    case 54
         column_letter = 'BB'
    case 55
        column_letter = 'BC'
    case 56
        column_letter = 'BD'
    case 57
        column_letter = 'BE'
    case 58 
        column_letter = 'BF'
    case 59 
        column_letter = 'BG'
    case 60
        column_letter = 'BH'
    case 61 
        column_letter = 'BI'
    case 62
        column_letter = 'BJ'
    case 63
        column_letter = 'BK'
    case 64
        column_letter = 'BL'
    case 65
        column_letter = 'BM'  
    case 66
        column_letter = 'BN'
    case 67 
        column_letter = 'BO'
    case 68 
        column_letter = 'BP'
    case 69 
        column_letter = 'BQ'
    case 70 
        column_letter = 'BR'
    case 71
        column_letter = 'BS'
    case 72
        column_letter = 'BT'
    case 73 
        column_letter = 'BU'
    case 74
        column_letter = 'BV'
    case 75
        column_letter = 'BW'
    case 76
        column_letter = 'BX'
    case 77
        column_letter = 'BY'
    case 78
        column_letter = 'BZ'
end