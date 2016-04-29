% Available Handles
    % handles.output = hObject;
    % handles.oldfilepath = pwd;
    % handles.filename = 'nofile';
    % handles.pathname = 'nopath';
    % handles.home_dir = pwd;
    % handles.worksheet = 'noworksheet';
    % handles.column_number = '1';
    % handles.column_letter = 'A';

% setting up files for reading and writing to excel
pathname = handles.pathname
filename = handles.filename
home_dir = handles.home_dir
worksheet = handles.worksheet
worksheet2 = strcat (worksheet, ' Last 5')

cd (pathname)

rows_cols = size(save_data)          % Gets array size (rows,columns)
max_row = rows_cols (1,1)       % Separates into row and col variables
max_col = rows_cols (1,2)

save_data2 = save_data
save_data = {'blank', 0}
exp_name = {'blank'}

current_col = 1
current_row = 1
save_row = 1

while current_col < max_col + 1         
    
    % Gets header data and stores in Column 1
    exp_name = save_header (1, current_col)
    save_data (current_col, 1) = exp_name
       
    current_row = 180
    last_five = 0
    i = 1
    
    
    % Add up the normalized slopes
    for current_row = 180:210
        last_five = last_five + save_data2 (current_row, current_col)
        i = i + 1
        current_row = current_row + 1
    end
    
    % Calculate % ltp of last 5 minutes
    perc_ltp = last_five / i
    
    % Set up data for saving
    save_data (save_row, 2) = num2cell (perc_ltp)
    save_row = save_row + 1
    current_col = current_col + 1
end

% save to the new worksheet
os = ispc

if os ==1
    xlswrite(filename, save_data, worksheet2)
    status = strcat ('Analysis complete. Data saved as ', filename)
    
else
    cd (home_dir)
    xlwrite(filename2, save_data, worksheet2)
    status = strcat ('Analysis complete. Data saved as ', filename2)
end
set(handles.txt_status,'String',status)

cd (home_dir)