folder_path = pwd; % replace with your folder path
[~, folderName, ~] = fileparts(folder_path);
OutputExcelName = strcat(folderName,' (comments)','.xlsx');
output_file = OutputExcelName; % replace with your desired output file name

files = dir(fullfile(folder_path, '*.par'));

output = cell(length(files), 4);
output(1,:) = {'Image', 'Comment', 'Date', 'Time'};

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
    % Convert to datetime object
    dt = datetime(dt_str, 'InputFormat', 'dd.MM.yyyy HH:mm');

    % Extract date and time components
    date_str = datestr(dt, 2);
    time_str = datestr(dt, 15);

    image_str = strtok(files(i).name, '_');
    
    output(i+1,:) = {image_str, comment_text, date_str, time_str};
end

xlswrite(fullfile(folder_path, output_file), output);
