% -----------------------------------------------------------------------%
% Author: Haley E. Speed, PhD                                            %
% Department of Neurology                                                %
% University of Texas Southwestern Medical Center                        %
% Dallas, TX                                                             %   
%-------------------------------------------------------------------------

% Analysis of raw Paired Pulse experiment data 
% Must be run by the parent program "Ephys Analysis_OSX.m
% This routine was written specifically for our PPR protocol, though future
% versions will be more accommodating of other PPR protocols
% See help file for more info

% set text box status to 'Working'
analyze_status = 'Working';

%Read in the directory, filename, and worksheet selected from the GUI
pathname = handles.pathname;
filename = handles.filename;
worksheet = handles.worksheet;
home_dir = handles.home_dir;
summary_data = 0;

%Change to the directory where the xlsx file is located
cd (pathname);

% Read the data from the excel worksheet into a variable 
[data,header] = xlsread (filename,worksheet);

% Initialize variables
rows_cols = size(data);          % Gets array size (rows,columns). Ignores half-empty columns. No idea why
max_row = rows_cols (1,1);       % Separates into row and col variables
max_col = rows_cols (1,2);
current_col = 1 ;                % Columns are slope 1 or slope 2 measurements
current_row = 1;                 % Rows correspond to sweep #
worksheet2 = strcat(worksheet,' Analyzed');
save_header = ' ';
save_data = 0;
save_col = 0;
save_header = {'blank'};
save_size = 0;
%Checks to see if the excel spreadsheet is set up correctly
file_check = rem(max_col,7);

if file_check == 0
    analyze_status = 'File is setup properly';
    set(handles.txt_status,'String',analyze_status);

    % Gets the party started
    while current_col < max_col

        % Need to determine if this is an ISI column
        col_check = rem(current_col-1,7);
    
        % is this the first column of a data set?
        if col_check == 0 | current_col == 1;
            
            % Setting up the header
            % Header is a Cell data type, so strings/numbers have to be
            % between brackets to indicate a cell data type
            exp_name = header (1, current_col);
            
            % The column math is different for the first set because the
            % last column in the worksheet is 0. 
            if current_col == 1
                save_header(end) = exp_name;
                save_header(end + 1) = {'Slope 1'};
                save_header(end + 1) = {'Slope 2'};
                save_header(end + 1) = {'PPR'};
            % For the second set of data, the last column in the worksheet
            % is 4. If you don't add + 1, you overwrite column 4.
            else
                save_header(end + 1) = exp_name;
                save_header(end + 1) = {'Slope 1'};
                save_header(end + 1) = {'Slope 2'};
                save_header(end + 1) = {'PPR'};
            end
        end
        % Advances to the next column, but no math is done until the next
        % set of data is reached
        current_col = current_col + 1; 
    end
    
    % Reading in the data and doing math on the raw data
    current_col = 1;
    current_row = 1;
    save_col = 0;
    save_row = 1;

    while current_col < 16 %max_col
        col_check = rem(current_col-1,7);
    
        if col_check == 0 | current_col == 1  
                        
            % Build an average slope 1, slope 2, and PPR for each ISI
            isi_num = 1;
            slope1_30 = 0;
            slope2_30 = 0;
            slope1_50 = 0;
            slope2_50 = 0;
            slope1_80 = 0;
            slope2_80 = 0;
            slope1_100 = 0;
            slope2_100 = 0;
            slope1_200 = 0;
            slope2_200 = 0;
            slope1_500 = 0;
            slope2_500 = 0;
            
            % Keep a tally of the sample sizes for each ISI
            sample_size30 = 0;
            sample_size50 = 0;
            sample_size80 = 0;
            sample_size100 = 0;
            sample_size200 = 0;
            sample_size500 = 0;
            
            % Passes through all rows and adds the slope 1 and slope 2
            % values to the appropriate ISI running tally. In this case,
            % sweep 1 = 50, sweep 2 = 500, sweep 3 = 100, sweep 4 = 80,
            % sweep 5 = 30 and sweep 6 = 200 ms. For example, Slope 2 
            % corresponds to the 50 ms ISI every 7 rows {sweeps) 
            for current_row = 1:max_row
                
                switch isi_num
                    case 1  % Row 1 Column 1 = Slope 1, Row 1 Column 3 = Slope 2
                        slope1_50 = slope1_50 + data (current_row, current_col);
                        slope2_50 = slope2_50 + data (current_row, current_col + 2);
                        sample_size50 = sample_size50 + 1;
                        isi_num = isi_num + 1;
                    
                    case 2   % Row 2 Column 1 = Slope 1, Row 2 Column 7 = Slope 2
                        slope1_500 = slope1_500 + data (current_row, current_col);
                        slope2_500 = slope2_500 + data (current_row, current_col + 6);
                        sample_size500 = sample_size500 +1
                        isi_num = isi_num + 1
                    
                    case 3  % Row 3 Column 1 = Slope 1, Row 3 Column 5 = Slope 2
                        slope1_100 = slope1_100 + data (current_row, current_col);
                        slope2_100 = slope2_100 + data (current_row, current_col + 4);
                        sample_size100 = sample_size100 + 1;
                        isi_num = isi_num + 1;
                   
                    case 4  % Row 4 Column 1 = Slope 1, Row 4 Column 4 = Slope 2 
                        slope1_80 = slope1_80 + data (current_row, current_col);
                        slope2_80 = slope2_80 + data (current_row, current_col + 3);
                        sample_size80 = sample_size80 + 1;
                        isi_num = isi_num + 1;
                    
                    case 5 % Row 5 Column 1 = Slope 1, Row 5 Column 2 = Slope 2
                        slope1_30 = slope1_30 + data (current_row, current_col);
                        slope2_30 = slope2_30 + data (current_row, current_col + 1);
                        sample_size30 = sample_size30 + 1;
                        isi_num = isi_num + 1;
                    
                    case 6 % Row 6 Column 1 = Slope 1, Row 6 Column 6 = Slope 2
                        slope1_200 = slope1_200 + data (current_row, current_col);
                        slope2_200 = slope2_200 + data (current_row, current_col + 5);
                        sample_size200 = sample_size200 + 1;
                        
                        % Calculate averages and enter into a new array to
                        % be saved in the excel workbook as '[Sheetname] Analyzed'
                        % One row per ISI with four columns per row --> One 
                        % average slope 1, one average slope 2, and one average 
                        % PPR per ISI
                        if current_col == 1 % First dataset has to be treated differently
                            save_col = 1;
                            save_data (1, save_col) = 30; % Empty worksheet, so can start at row 1, column 1
                            save_data (1, save_col + 1) = slope1_30/sample_size30;
                            save_data (1, save_col + 2) = slope2_30/sample_size30;
                            save_data (1, save_col + 3) = save_data (1, save_col + 2) / save_data (1, save_col + 1);
                        
                            save_data (2, save_col) = 50;
                            save_data (2, save_col + 1) = slope1_50/sample_size50;
                            save_data (2, save_col + 2) = slope2_50/sample_size50;
                            save_data (2, save_col + 3) = save_data (2, save_col + 2) / save_data (2, save_col + 1);
                       
                            save_data (3, save_col) = 80;
                            save_data (3, save_col + 1) = slope1_80/sample_size80;
                            save_data (3, save_col + 2) = slope2_80/sample_size80;
                            save_data (3, save_col + 3) = save_data (3, save_col + 2) / save_data (3, save_col + 1);
                        
                            save_data (4, save_col) = 100;
                            save_data (4, save_col + 1) = slope1_100/sample_size100;
                            save_data (4, save_col + 2) = slope2_100/sample_size100;
                            save_data (4, save_col + 3) = save_data (4, save_col + 2) / save_data (4, save_col + 1);
                        
                            save_data (5, save_col) = 200;
                            save_data (5, save_col + 1) = slope1_200/sample_size200;
                            save_data (5, save_col + 2) = slope2_200/sample_size200;
                            save_data (5, save_col + 3) = save_data (5, save_col + 2)/ save_data (5, save_col + 1);
                        
                            save_data (6, save_col) = 500;
                            save_data (6, save_col + 1) = slope1_500/sample_size500;
                            save_data (6, save_col + 2) = slope2_500/sample_size500;
                            save_data (6, save_col + 3) = save_data (6, save_col + 2) / save_data (6, save_col + 1);
                        
                        % For all subsequent PPR datasets, the starting
                        % column has to be after those already in the save
                        % array. The variable save_col keeps track of how
                        % many sets have been analyzed before.
                        else                     
                            save_data (1, save_col + 1) = 30;
                            save_data (1, save_col + 2) = slope1_30/sample_size30;
                            save_data (1, save_col + 3) = slope2_30/sample_size30;
                            save_data (1, save_col + 4) = save_data (1, save_col + 3) / save_data (1, save_col + 2);
                        
                            save_data (2, save_col + 1) = 50;
                            save_data (2, save_col + 2) = slope1_50/sample_size50;
                            save_data (2, save_col + 3) = slope2_50/sample_size50;
                            save_data (2, save_col + 4) = save_data (2, save_col + 3) / save_data (2, save_col + 2);
                       
                            save_data (3, save_col + 1) = 80;
                            save_data (3, save_col + 2) = slope1_80/sample_size80;
                            save_data (3, save_col + 3) = slope2_80/sample_size80;
                            save_data (3, save_col + 4) = save_data (3, save_col + 3) / save_data (3, save_col + 2);
                        
                            save_data (4, save_col + 1) = 100;
                            save_data (4, save_col + 2) = slope1_100/sample_size100;
                            save_data (4, save_col + 3) = slope2_100/sample_size100;
                            save_data (4, save_col + 4) = save_data (4, save_col + 3) / save_data (4, save_col + 2);
                        
                            save_data (5, save_col + 1) = 200;
                            save_data (5, save_col + 2) = slope1_200/sample_size200;
                            save_data (5, save_col + 3) = slope2_200/sample_size200;
                            save_data (5, save_col + 4) = save_data (5, save_col + 3)/ save_data (5, save_col + 2);
                        
                            save_data (6, save_col + 1) = 500;
                            save_data (6, save_col + 2) = slope1_500/sample_size500;
                            save_data (6, save_col + 3) = slope2_500/sample_size500;
                            save_data (6, save_col + 4) = save_data (6, save_col + 3) / save_data (6, save_col + 2);
                        end
                        isi_num = 1;  % Resets counter for ISI/row
                end
                current_row = current_row + 1; % Go to the next row
            end
        end
        % When no more rows are left in the column, go to the next column
        save_rowscols = size (save_data); % Gets size of save_data array
        
        % Determines which is the last column in the save array as the 
        % starting point for saving the analyzed data in the next dataset
        save_col = save_rowscols (1,2);  
        current_col = current_col + 1;
    end
    
    
    % save the data using the excelsave routine. Works on PC or Mac_OS
    % Yosimite. Not supported beyond Yosimite
    cd (home_dir)
    excelsave
    save_header
else
    error_text = ['File does not seem to be setup properly.'...
        ' Make sure data are arranged in three columns: 1) Interstimulus'... 
        'Interval 2) Slope 1 3) Slope 2.'];
    set(handles.txt_status,'String', error_text);
end


