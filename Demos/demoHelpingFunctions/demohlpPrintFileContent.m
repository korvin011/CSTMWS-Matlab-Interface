function demohlpPrintFileContent(FileName, NLines)
    Txt = fileread(FileName);
    Lines = strsplit(Txt,newline).'; % split on lines
    if NLines>=length(Lines), 
        NLines = length(Lines);
        EoFPrefix = '';
    else
        EoFPrefix = '...\n';
    end
    fprintf('--------- Content of "%s" -----------\n',FileName);
    fprintf('%s',Lines{1:NLines});
    fprintf([EoFPrefix '-------------------------------------------------\n']);
end