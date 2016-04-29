
% -----------------------------------------------------------------------%
% Author: Haley E. Speed, PhD                                            %
% Department of Neurology                                                %
% University of Texas Southwestern Medical Center                        %
% Dallas, TX                                                             %   
%-------------------------------------------------------------------------

% Routine for analyzing raw slope/peak data from input/output curves and
% assumes that there are 5 sweeps per stimulus intensity (i.e. 5 rows per
% stimulus intensity). It finds the mean at each intensity.

% Available Handles
    % handles.output = hObject;
    % handles.oldfilepath = pwd;
    % handles.filename = 'nofile';
    % handles.pathname = 'nopath';
    % handles.worksheet = 'noworksheet';
    % handles.column_number = '1';
    % handles.column_letter = 'A';

% setting up files for reading and writing to excel
pathname = handles.pathname
filename = handles.filename
worksheet = handles.worksheet
home_dir = handles.home_dir
summary_data = 0

% Change directory to working folder
cd (pathname)

% read in data stored in worksheet and store it in an array (variable)
[data,header] = xlsread (filename,worksheet)

% Initialize variables
rows_cols = size(data)          % Gets array size (rows,columns)
max_row = rows_cols (1,1)       % Separates into row and col variables
max_col = rows_cols (1,2)
current_col = 1
current_row = 1
worksheet2 = strcat(worksheet,' Analyzed')
save_header = ' '
save_data = 0
save_col = 0
save_header = {'blank'}
exp_name = {}



% Gets header data
    while current_col < max_col + 1
        % Save first column with exp_name
            
        exp_name = header (1, current_col)
      
        if current_col == 1
            save_header(end) = exp_name
        else
            save_header(end + 1) = exp_name   
        end
 
        current_col = current_col + 1
    end
    
% Reset row and column to 1 for I/O curve analysis    
current_col = 1
current_row = 1
quotient = (max_row)/5

% Search every 5 data points and add to array
while current_col < max_col + 1
    save_row = 1
    current_row = 1
    avg_row = 1
    sum = 0
    slope = 0
    while current_row < max_row + 1
    
        % Average every 5 data points and add to array
        if current_row < avg_row + 5
            slope = data(current_row, current_col)
            sum = sum + slope
            current_row = current_row + 1
        else
            average = sum/5
            save_data(save_row,current_col) = average
            avg_row = avg_row + 5
            save_row = save_row + 1
            sum = 0
            slope = 0
        end
    end
    if current_row == max_row + 1
        average = sum/5
        save_data(save_row,current_col) = average
        avg_row = avg_row + 5
        save_row = save_row + 1
        sum = 0
        slope = 0
    end    
    current_col = current_col + 1
end

% change to the home directory
cd (home_dir)

% Write the array to the new worksheet
excelsave



