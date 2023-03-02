folder_path = pwd; % replace with your folder path
[~, folderName, ~] = fileparts(folder_path);
OutputExcelName = strcat(folderName,' (scan conditions and comments)','.xlsx');
output_file = OutputExcelName; % replace with your desired output file name

files = dir(fullfile(folder_path, '*.par'));

output = cell(length(files), 5);
output(1,:) = {'Image', 'V;I', 'Comment', 'Date', 'Time'};

for i = 1:length(files)
    file_path = fullfile(folder_path, files(i).name);
    file_contents = fileread(file_path);
    lines = splitlines(file_contents);
    
    comment_line = find(startsWith(lines, 'Comment                         :'), 1);
    if isempty(comment_line)
        comment_line = find(startsWith(lines, 'Comment                          :'), 1);
    end
    comment_text = extractAfter(lines{comment_line}, 'Comment                         : ');
    
    date_line = find(startsWith(lines, 'Date                            :'), 1);
    dt_str = extractAfter(lines{date_line}, 'Date                            : ');
    
    V_line = find(startsWith(lines, 'Gap Voltage                     :'), 1);
    V_str = extractAfter(lines{V_line}, 'Gap Voltage                     : ');
    
    I_line = find(startsWith(lines, 'Feedback Set                    :'), 1);
    I_str = extractAfter(lines{I_line}, 'Feedback Set                    : ');
    
    V_match = regexp(V_str, '-?\d+\.?\d*', 'match'); % extract number as a string
    V_num = str2double(V_match{1}); % convert string to number
    rounded_V = round(V_num, 2); % round to 2 digits after the decimal point
    V_str = num2str(rounded_V);

    I_match = regexp(I_str, '-?\d+\.?\d*', 'match'); % extract number as a string
    I_num = str2double(I_match{1}); % convert string to number
    rounded_I = round(I_num, 2); % round to 2 digits after the decimal point
    I_str = num2str(rounded_I);
        
    % Convert to datetime object
    dt = datetime(dt_str, 'InputFormat', 'dd.MM.yyyy HH:mm');

    % Extract date and time components
    date_str = datestr(dt, 2);
    time_str = datestr(dt, 15);

    image_str = strtok(files(i).name, '_');
    
    output(i+1,:) = {image_str, strcat(V_str,' V;',I_str,' nA'), comment_text, date_str, time_str};
end

xlswrite(fullfile(folder_path, output_file), output);
