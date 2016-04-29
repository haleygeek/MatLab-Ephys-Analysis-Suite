% -----------------------------------------------------------------------%
% Author: Haley E. Speed, PhD                                            %
% Department of Neurology                                                %
% University of Texas Southwestern Medical Center                        %
% Dallas, TX                                                             %   
%-------------------------------------------------------------------------

% Routine takes raw slope data (i.e. from Clampfit), the point at which you
% indicate DHPG hits the slice (onset of acute depression) and calculates a
% 10 minute wash on, baseline, and post-wash % increase or decrease
% relative to baseline. Future versions will allow for variable wash-in
% times.

% setting up files for reading and writing to excel
pathname = handles.pathname;
filename = handles.filename;
home_dir = handles.home_dir;
worksheet = handles.worksheet;
dhpg_on = handles.dhpg_on
wash = handles.wash

% Setting up boundaries within the raw data
baseline = dhpg_on - 30 % 10 minutes before onset of DHPG at the slice
washout = dhpg_on + ((wash * 60)/20) % user defined minutes after onset of DHPG
ltd_start = washout + 165
ltd_end = washout + 180

cd (pathname)

% read in data stored in worksheet and store it in an array (variable)
[data,header] = xlsread (filename,worksheet)

% Initialize variables
rows_cols = size(data);          % Gets array size (rows,columns)
max_row = rows_cols (1,1);       % Separates into row and col variables
max_col = rows_cols (1,2);
current_col = 1;
current_row = 1;

worksheet2 = strcat(worksheet,' Analyzed');
worksheet3 = strcat(worksheet, ' Summary');

save_ltd = {};
save_data = 0;
save_col = 0;
save_summary = {};

save_ltd = {'blank'};
exp_name = {'blank'};

save_summary (1, 1) = cellstr('Experiment');
save_summary (1,2) = cellstr('% Basline @ 55-60 min');

% Save first column of normalized data and summarized data with exp_name
while current_col <= max_col
    current_row == 1
    exp_name = header (1, current_col);
    save_ltd (1, current_col) = cellstr(exp_name);
    save_summary (current_col + 1, 1) = cellstr(exp_name);
    current_col = current_col + 1
end

col = 1;
sum = 0;
average = 0;

% Main calculations 
while col < max_col + 1;
    
    % Average the last 10 minutes of baseline
    
    row = baseline;
    i = 1;
    
    for row = baseline:dhpg_on;
        sum = sum + data (row, col);
        i = i+1;
    end
    average = sum / i;
        
    % Divide all rows by the average and normalize to 100%
    last_five = 0
    row = 1;
    j = 1;
    while row < max_row + 1;
        save_data (row, col) = (data (row, col) / average) * 100;
        save_ltd (row + 1, col) = num2cell(save_data (row, col));
        
        % If the row is in the last 5 minutes, save it for the average
        if row >= ltd_start & row <= ltd_end;
            last_five = last_five + save_data (row, col)
            j = j + 1;
        end
        
        row = row + 1;
    end

    % Calculate % LTD of last 5 minutes
    perc_ltd = last_five / j;
    save_summary (col + 1, 2) = num2cell(perc_ltd)
    
    % Reset internal variables and start a new column
    sum = 0;
    average = 0;
    col = col + 1;
end 

% Reset the current directory to call the next scripts
cd (home_dir)

% Save the worksheets
if ispc == 1
    cd (pathname)
    xlswrite (filename, save_ltd, worksheet2);
    xlswrite (filename, save_summary, worksheet3);

    analyze_status = strcat ('Analysis complete. ', worksheet, ' and ', worksheet2, ' have been added to: ',pathname, filename);
    cd (home_dir);
else
    % Importing java files for saving xlsx on Macs
    javaaddpath('/library/java/extensions/jxl.jar');
    javaaddpath('/library/java/extensions/MXL.jar');

    import mymxl.*;
    import jxl.*; 
   
    % Have to save to a different file or else get java error 
    pathname2 = strcat (pathname, 'Analyzed/')
    filename2 = strcat(pathname2, filename, ' Analyzed')
    
    if ~exist(pathname2, 'dir')
        mkdir(pathname2)
    end
    
    % Save to the new worksheet with a java workaround
    xlwrite(filename2, save_ltd, worksheet2);
    xlwrite (filename2, save_summary, worksheet3);
   
    % Display save path in text_box
    analyze_status = strcat ('Analysis complete. Data saved as ', filename2)
    
end

% Update text box with path name
set(handles.txt_status,'String', analyze_status)

% Reset Folder
cd(home_dir)
