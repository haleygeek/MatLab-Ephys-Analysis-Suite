% set text box status to 'Working'
analyze_status = 'Working'

%Read in the directory, filename, and worksheet selected from the GUI
pathname = handles.pathname
filename = handles.filename
worksheet = handles.worksheet
home_dir = handles.home_dir
summary_data = 0

%Change to the directory where the xlsx file is located
cd (pathname)

% Read the data from the excel worksheet into a variable 
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
ppr_header = cellstr ('Slope1/Slope2')

% Check for proper file setup (3 columns/experiment)
file_check = rem(max_col,3)
if file_check == 0
    analyze_status = 'File is setup properly'
    set(handles.txt_status,'String',analyze_status);


    %Starts Analysis for one column
    while current_col < max_col

        % Need to determine whether this is a ISI column, pulse 1 column, or pulse 2 column
        col_sub = (current_col-1)
        col_check = rem(col_sub,3)
    
    
        % Save first column with exp_name
        % Save second column with 'PPR'
        if col_check == 0 | current_col ==1
            disp ('Start Loop----------------------------------------------')

            exp_name = header (1, current_col)
           % str_name = exp_name
        
            if current_col == 1
                save_header(end) = exp_name
                save_header(end + 1) = ppr_header
            
            else
                save_header(end + 1) = exp_name 
                save_header(end + 1) = ppr_header
            end
            disp ('end loop------------------------------------------------')
        else
        
        
            current_col = current_col + 1
        end
        current_col = current_col + 1
    end

    current_col = 1
    current_row = 1

    while current_col < max_col
    
        col_sub = (current_col-1)
        col_check = rem(col_sub,3)
    
        if col_check == 0 | current_col == 1  
            disp ('Start Loop----------------------------------------------')
            save_size = size (save_data)
            last_col = save_size (1,2) 
        
            for current_row = 1:max_row
                interval = data(current_row,current_col)
                slope1 = data (current_row, current_col + 1)
                slope2 = data (current_row, current_col + 2)
                ratio = slope2 / slope1
            
            
        
                if current_col == 1
                    save_data(current_row, 1) = interval
                    save_data(current_row, 2) = ratio
               
                else
                    save_data(current_row, last_col + 1) = interval
                    save_data(current_row, last_col + 2) = ratio
                
                end
                current_row = current_row + 1
        
            end
            disp ('end loop------------------------------------------------')
        else
            analyze_status = 'Error. Check the file name and format.'
        end
        current_col = current_col + 1
    end
    
    % save the data
    cd (home_dir)
    excelsave
else
    error_text = ['File does not seem to be setup properly.'...
        ' Make sure data are arranged in three columns: 1) Interstimulus'... 
        'Interval 2) Slope 1 3) Slope 2.']
    set(handles.txt_status,'String', error_text)
end


