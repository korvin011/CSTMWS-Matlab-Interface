classdef TCSTInterface < handle
% @TITLE ~-- CST Microwave Studio to MATLAB interface
% 
% @DESCRIPTION
%   The class allows to call various functionality of the CST Microwave Studio, including
% manipulating with parameters in the existing project, running solver and obtaining
% solution results such as S-parameters, far-field pattern and other.
%   Additionally it allows calculate the cost function for the CST optimizer in MATLAB.
%   For the detailed description of the class methods and properties, please see the demos
% coming with this class.
% 
% @AUTHOR    Oleg Iupikov <oleg.iupikov@chalmers.se, lichne@gmail.com>
% @VERSION   2.2.000
% @FIRSTEDIT 15-01-2018
% @LASTEDIT  17-07-2019
    
    properties (Constant, Access = private)
%         SETTINGS_FILE = fullfile(prefdir,'CSTInterface.mat');
        FVersion = '2.2.000'  % Interface version
        FCSTVersion = '2019'  % Tested latest CST version
        FAuthor = 'Oleg Iupikov, <a href="mailto:lichne@gmail.com">lichne@gmail.com</a>, <a href="mailto:oleg.iupikov@chalmers.se">oleg.iupikov@chalmers.se</a>';
        FOrganization = 'Chalmers University of Technology, Sweden';
        FFirstEdit = '15-01-2018';
        FLastEdit  = '17-07-2019';
    end
    
    properties (Access = private)
        FApp
        FProj
        FSolver
        FCalculateYZMatrices = [] % true - activate automatic calculation;  false - deactivate;  [] - do not change the current setting
        FSilentMode = false
        FNearestOrInterpolated = 'Nearest'
    end
    
    properties (Dependent)
        % Can be \code{true}, \code{false} or \code{[]}. Activates/deactivated the CST
        % postprocessing option under ``\code{Post-processing/S-parameter calculation/Always
        % Calculate Z and Y Matrices}'' for the currently connected project. Assignment
        % \code{[]} doesn't do anything.
        CalculateYZMatrices
        
        % Reference to the CST Application COM object. Assigned by the method
        % \refmethod{ConnectToCSTOrStartIt}.
        Application
        
        % Reference to the CST Project COM object. Assigned by the method
        % \refmethod{OpenProject}.
        Project
        
        SilentMode
        
        % For many result-retrieving methods we can specify points (e.g.
        % frequencies) for which we want to get the result. However, not
        % always specified point will exist in the obtained result. This
        % option defines how to deal with such situation. 
        % When \code{NearestOrInterpolated='Exact'}, the funtion value must
        % exist at the requested point, otherwise an error will be thrown. 
        % When \code{NearestOrInterpolated='Nearest'} (default), the function value will
        % be returned for the point closest to the requested one. 
        % When \code{NearestOrInterpolated='Interpolated'}, a linear interpolation
        % will be performed to get the function value at the requested
        % point.
        NearestOrInterpolated
    end
    
    % =====================================================================
    % ==================== GET and SET methods ============================
    % =====================================================================
    methods
        % -----------------------------------------------------------------
        % CalculateYZMatrices
        % -----------------------------------------------------------------
        function data = get.CalculateYZMatrices(this)
            data = this.FCalculateYZMatrices;
        end
        function set.CalculateYZMatrices(this, data)
            assert(islogical(data) || isempty(data), '"CalculateYZMatrices" must be "true", "false" or "[]" (means "do not change").');
            this.FCalculateYZMatrices = data;
            this.setCalculateYZMatrices(data);
        end
        % -----------------------------------------------------------------
        % Application
        % -----------------------------------------------------------------
        function data = get.Application(this)
            data = this.FApp;
        end
        function set.Application(~,~)
            error('"Application" is read-only property. Use "ConnectToCSTOrStartIt()" method to assign this property.');
        end
        % -----------------------------------------------------------------
        % Project
        % -----------------------------------------------------------------
        function data = get.Project(this)
            data = this.FProj;
        end
        function set.Project(~,~)
            error('"Project" is read-only property. Use "OpenProject(...)" method to assign this property.');
        end
        % -----------------------------------------------------------------
        % SilentMode
        % -----------------------------------------------------------------
        function data = get.SilentMode(this)
            data = this.FSilentMode;
        end
        function set.SilentMode(this, data)
            validateattributes(data,{'logical','numeric'},{'nonempty','scalar','nonnegative'});
            this.FSilentMode = data;
        end
        % -----------------------------------------------------------------
        % NearestOrInterpolated
        % -----------------------------------------------------------------
        function data = get.NearestOrInterpolated(this)
            data = this.FNearestOrInterpolated;
        end
        function set.NearestOrInterpolated(this, data)
            ValidStrings = {'Exact', 'Nearest', 'Interpolated'};
            data = validatestring(data, ValidStrings, '', 'NearestOrInterpolated');
            this.FNearestOrInterpolated = data;
        end
%         % -----------------------------------------------------------------
%         % 
%         % -----------------------------------------------------------------
%         function data = get.(this)
%             data = this.F;
%         end
%         function set.(this, data)
%             this.F = data;
%         end
    end
     
    % =====================================================================
    % ======================== STATIC Methods ============================
    % =====================================================================
    methods (Static)
        function File = GetFullPath(File, Style)
        % @ TITLE ~-- Get absolute canonical path of a file or folder
        % 
        % @DESCRIPTION
        %   GetFullPath by Jan Simon
        % (\href{https://se.mathworks.com/matlabcentral/fileexchange/28249-getfullpath}{https://se.mathworks.com/matlabcentral/fileexchange/28249-getfullpath})
        % 
        %   Absolute path names are safer than relative paths, when e.g. a GUI or TIMER
        % callback changes the current directory. Only canonical paths without "." and
        % ".." can be recognized uniquely.
        % Long path names (>259 characters) require a magic initial key \code{'\\?\'} to be
        % handled by Windows API functions, e.g. for Matlab's FOPEN, DIR and EXIST.
        % 
        % @SYNTAX
        % 
        % FullName = TCSTInterface.GetFullPath(Name)
        % FullName = TCSTInterface.GetFullPath(Name, Style)
        % 
        % @EXAMPLES
        % 
        %   cd(tempdir);                                    % Assumed as 'C:\Temp' here
        %   TCSTInterface.GetFullPath('File.Ext')           % 'C:\Temp\File.Ext'
        %   TCSTInterface.GetFullPath('..\File.Ext')        % 'C:\File.Ext'
        %   TCSTInterface.GetFullPath('..\..\File.Ext')     % 'C:\File.Ext'
        %   TCSTInterface.GetFullPath('.\File.Ext')         % 'C:\Temp\File.Ext'
        %   TCSTInterface.GetFullPath('*.txt')              % 'C:\Temp\*.txt'
        %   TCSTInterface.GetFullPath('..')                 % 'C:\'
        %   TCSTInterface.GetFullPath('..\..\..')           % 'C:\'
        %   TCSTInterface.GetFullPath('Folder\')            % 'C:\Temp\Folder\'
        %   TCSTInterface.GetFullPath('D:\A\..\B')          % 'D:\B'
        %   TCSTInterface.GetFullPath('\\Server\Folder\Sub\..\File.ext')  
        %                                                   %'\\Server\Folder\File.ext'
        %   TCSTInterface.GetFullPath({'..', 'new'})        % {'C:\', 'C:\Temp\new'}
        %   TCSTInterface.GetFullPath('.', 'fat')           % '\\?\C:\Temp\File.Ext'

        % @ARGUMENTS <<< some bugs in LaTeX doc generator... not included to PDF help 
        % 
        % @ARG Name - String or cell string, absolute or relative name of a file or
        %      folder. The path need not exist. Unicode strings, UNC paths and long
        %      names are supported.
        % @ARG Style - Style of the output as string, optional, default: \code{'auto'}.
        %      Values can be: \code{'auto'}: Add \code{'\\?\'} or \code{'\\?\UNC\'} for long names
        %      on demand; \code{'lean'): Magic string is not added; \code{'fat'}:
        %      Magic string is added for short names also. The \refmetarg{Style} is
        %      ignored when not running under Windows. 
        %
        % @RETURN
        %
        % @RET FullName - Absolute canonical path name as string or cell string.
        %          For empty strings the current directory is replied.
        %          \code{'\\?\'} or \code{'\\?\UNC'} is added on demand.
        %
        %
        % @EXAMPLES:
        % 
        %   cd(tempdir);                    % Assumed as 'C:\Temp' here
        %   GetFullPath('File.Ext')         % 'C:\Temp\File.Ext'
        %   GetFullPath('..\File.Ext')      % 'C:\File.Ext'
        %   GetFullPath('..\..\File.Ext')   % 'C:\File.Ext'
        %   GetFullPath('.\File.Ext')       % 'C:\Temp\File.Ext'
        %   GetFullPath('*.txt')            % 'C:\Temp\*.txt'
        %   GetFullPath('..')               % 'C:\'
        %   GetFullPath('..\..\..')         % 'C:\'
        %   GetFullPath('Folder\')          % 'C:\Temp\Folder\'
        %   GetFullPath('D:\A\..\B')        % 'D:\B'
        %   GetFullPath('\\Server\Folder\Sub\..\File.ext')
        %                                   % '\\Server\Folder\File.ext'
        %   GetFullPath({'..', 'new'})      % {'C:\', 'C:\Temp\new'}
        %   GetFullPath('.', 'fat')         % '\\?\C:\Temp\File.Ext'
        %
        % @OTHER
        % 
        % NOTE: The M- and the MEX-version create the same results, the faster MEX
        %   function works under Windows only.
        %   Some functions of the Windows-API still do not support long file names.
        %   E.g. the Recycler and the Windows Explorer fail even with the magic '\\?\'
        %   prefix. Some functions of Matlab accept 260 characters (value of MAX_PATH),
        %   some at 259 already. Don't blame me.
        %   The 'fat' style is useful e.g. when Matlab's DIR command is called for a
        %   folder with les than 260 characters, but together with the file name this
        %   limit is exceeded. Then "dir(GetFullPath([folder, '\*.*], 'fat'))" helps.
        % COMPILE:
        %   Automatic: InstallMex GetFullPath.c uTest_GetFullPath
        %   Manual:    mex -O GetFullPath.c
        %   Download:  http://www.n-simon.de/mex
        % Run the unit-test uTest_GetFullPath after compiling.
        %
        % Tested: Matlab 6.5, 7.7, 7.8, 7.13, WinXP/32, Win7/64
        %         Compiler: LCC2.4/3.8, BCC5.5, OWC1.8, MSVC2008/2010
        % Assumed Compatibility: higher Matlab versions
        % Author: Jan Simon, Heidelberg, (C) 2009-2013 matlab.THISYEAR(a)nMINUSsimon.de
        %
        % See also: CD, FULLFILE, FILEPARTS.

        % $JRev: R-G V:032 Sum:7Xd/JS0+yfax Date:15-Jan-2013 01:06:12 $
        % $License: BSD (use/copy/change/redistribute on own risk, mention the author) $
        % $UnitTest: uTest_GetFullPath $
        % $File: Tools\GLFile\GetFullPath.m $
        % History:
        % 001: 20-Apr-2010 22:28, Successor of Rel2AbsPath.
        % 010: 27-Jul-2008 21:59, Consider leading separator in M-version also.
        % 011: 24-Jan-2011 12:11, Cell strings, '~File' under linux.
        %      Check of input types in the M-version.
        % 015: 31-Mar-2011 10:48, BUGFIX: Accept [] as input as in the Mex version.
        %      Thanks to Jiro Doke, who found this bug by running the test function for
        %      the M-version.
        % 020: 18-Oct-2011 00:57, BUGFIX: Linux version created bad results.
        %      Thanks to Daniel.
        % 024: 10-Dec-2011 14:00, Care for long names under Windows in M-version.
        %      Improved the unittest function for Linux. Thanks to Paul Sexton.
        % 025: 09-Aug-2012 14:00, In MEX: Paths starting with "\\" can be non-UNC.
        %      The former version treated "\\?\C:\<longpath>\file" as UNC path and
        %      replied "\\?\UNC\?\C:\<longpath>\file".
        % 032: 12-Jan-2013 21:16, 'auto', 'lean' and 'fat' style.

        % Initialize: ==================================================================
        % Do the work: =================================================================

        % #############################################
        % ### USE THE MUCH FASTER MEX ON WINDOWS!!! ###
        % #############################################

        % Difference between M- and Mex-version:
        % - Mex does not work under MacOS/Unix.
        % - Mex calls Windows API function GetFullPath.
        % - Mex is much faster.

            % Magix prefix for long Windows names:
            if nargin < 2
               Style = 'auto';
            end

            % Handle cell strings:
            % NOTE: It is faster to create a function @cell\GetFullPath.m under Linux, but
            % under Windows this would shadow the fast C-Mex.
            if isa(File, 'cell')
               for iC = 1:numel(File)
                  File{iC} = TCSTInterface.GetFullPath(File{iC}, Style);
               end
               return;
            end

            % Check this once only:
            isWIN    = strncmpi(computer, 'PC', 2);
            MAX_PATH = 260;

            % Warn once per session (disable this under Linux/MacOS):
            persistent hasDataRead
            if isempty(hasDataRead)
               % Test this once only - there is no relation to the existence of DATAREAD!
               %if isWIN
               %   Show a warning, if the slower Matlab version is used - commented, because
               %   this is not a problem and it might be even useful when the MEX-folder is
               %   not inlcuded in the path yet.
               %   warning('JSimon:GetFullPath:NoMex', ...
               %      ['GetFullPath: Using slow Matlab-version instead of fast Mex.', ...
               %       char(10), 'Compile: InstallMex GetFullPath.c']);
               %end

               % DATAREAD is deprecated in 2011b, but still available. In Matlab 6.5, REGEXP
               % does not know the 'split' command, therefore DATAREAD is preferred:
               hasDataRead = ~isempty(which('dataread'));
            end

            if isempty(File)  % Accept empty matrix as input:
               if ischar(File) || isnumeric(File)
                  File = cd;
                  return;
               else
                  error(['JSimon:', mfilename, ':BadTypeInput1'], ...
                     ['*** ', mfilename, ': Input must be a string or cell string']);
               end
            end

            if ischar(File) == 0  % Non-empty inputs must be strings
               error(['JSimon:', mfilename, ':BadTypeInput1'], ...
                  ['*** ', mfilename, ': Input must be a string or cell string']);
            end

            if isWIN  % Windows: --------------------------------------------------------
               FSep = '\';
               File = strrep(File, '/', FSep);

               % Remove the magic key on demand, it is appended finally again:
               if strncmp(File, '\\?\', 4)
                  if strncmpi(File, '\\?\UNC\', 8)
                     File = ['\', File(7:length(File))];  % Two leading backslashes!
                  else
                     File = File(5:length(File));
                  end
               end

               isUNC   = strncmp(File, '\\', 2);
               FileLen = length(File);
               if isUNC == 0                        % File is not a UNC path
                  % Leading file separator means relative to current drive or base folder:
                  ThePath = cd;
                  if File(1) == FSep
                     if strncmp(ThePath, '\\', 2)   % Current directory is a UNC path
                        sepInd  = strfind(ThePath, '\');
                        ThePath = ThePath(1:sepInd(4));
                     else
                        ThePath = ThePath(1:3);     % Drive letter only
                     end
                  end

                  if FileLen < 2 || File(2) ~= ':'  % Does not start with drive letter
                     if ThePath(length(ThePath)) ~= FSep
                        if File(1) ~= FSep
                           File = [ThePath, FSep, File];
                        else                        % File starts with separator:
                           File = [ThePath, File];
                        end
                     else                           % Current path ends with separator:
                        if File(1) ~= FSep
                           File = [ThePath, File];
                        else                        % File starts with separator:
                           ThePath(length(ThePath)) = [];
                           File = [ThePath, File];
                        end
                     end

                  elseif FileLen == 2 && File(2) == ':'   % "C:" current directory on C!
                     % "C:" is the current directory on the C-disk, even if the current
                     % directory is on another disk! This was ignored in Matlab 6.5, but
                     % modern versions considers this strange behaviour.
                     if strncmpi(ThePath, File, 2)
                        File = ThePath;
                     else
                        try
                           File = cd(cd(File));
                        catch    % No MException to support Matlab6.5...
                           if exist(File, 'dir')  % No idea what could cause an error then!
                              rethrow(lasterror);
                           else  % Reply "K:\" for not existing disk:
                              File = [File, FSep];
                           end
                        end
                     end
                  end
               end

            else         % Linux, MacOS: ---------------------------------------------------
               FSep = '/';
               File = strrep(File, '\', FSep);

               if strcmp(File, '~') || strncmp(File, '~/', 2)  % Home directory:
                  HomeDir = getenv('HOME');
                  if ~isempty(HomeDir)
                     File(1) = [];
                     File    = [HomeDir, File];
                  end

               elseif strncmpi(File, FSep, 1) == 0
                  % Append relative path to current folder:
                  ThePath = cd;
                  if ThePath(length(ThePath)) == FSep
                     File = [ThePath, File];
                  else
                     File = [ThePath, FSep, File];
                  end
               end
            end

            % Care for "\." and "\.." - no efficient algorithm, but the fast Mex is
            % recommended at all!
            if ~isempty(strfind(File, [FSep, '.']))
               if isWIN
                  if strncmp(File, '\\', 2)  % UNC path
                     index = strfind(File, '\');
                     if length(index) < 4    % UNC path without separator after the folder:
                        return;
                     end
                     Drive            = File(1:index(4));
                     File(1:index(4)) = [];
                  else
                     Drive     = File(1:3);
                     File(1:3) = [];
                  end
               else  % Unix, MacOS:
                  isUNC   = false;
                  Drive   = FSep;
                  File(1) = [];
               end

               hasTrailFSep = (File(length(File)) == FSep);
               if hasTrailFSep
                  File(length(File)) = [];
               end

               if hasDataRead
                  if isWIN  % Need "\\" as separator:
                     C = dataread('string', File, '%s', 'delimiter', '\\');  %#ok<REMFF1>
                  else
                     C = dataread('string', File, '%s', 'delimiter', FSep);  %#ok<REMFF1>
                  end
               else  % Use the slower REGEXP, when DATAREAD is not available anymore:
                  C = regexp(File, FSep, 'split');
               end

               % Remove '\.\' directly without side effects:
               C(strcmp(C, '.')) = [];

               % Remove '\..' with the parent recursively:
               R = 1:length(C);
               for dd = reshape(find(strcmp(C, '..')), 1, [])
                  index    = find(R == dd);
                  R(index) = [];
                  if index > 1
                     R(index - 1) = [];
                  end
               end

               if isempty(R)
                  File = Drive;
                  if isUNC && ~hasTrailFSep
                     File(length(File)) = [];
                  end

               elseif isWIN
                  % If you have CStr2String, use the faster:
                  %   File = CStr2String(C(R), FSep, hasTrailFSep);
                  File = sprintf('%s\\', C{R});
                  if hasTrailFSep
                     File = [Drive, File];
                  else
                     File = [Drive, File(1:length(File) - 1)];
                  end

               else  % Unix:
                  File = [Drive, sprintf('%s/', C{R})];
                  if ~hasTrailFSep
                     File(length(File)) = [];
                  end
               end
            end

            % "Very" long names under Windows:
            if isWIN
               if ~ischar(Style)
                  error(['JSimon:', mfilename, ':BadTypeInput2'], ...
                     ['*** ', mfilename, ': Input must be a string or cell string']);
               end

               if (strncmpi(Style, 'a', 1) && length(File) >= MAX_PATH) || ...
                     strncmpi(Style, 'f', 1)
                  % Do not use [isUNC] here, because this concerns the input, which can
                  % '.\File', while the current directory is an UNC path.
                  if strncmp(File, '\\', 2)  % UNC path
                     File = ['\\?\UNC', File(2:end)];
                  else
                     File = ['\\?\', File];
                  end
               end
            end
        end 
       
        function [Field, Info] = ReadFarFieldSourceFile(FileOrDir, FileMask, NormalizationType)
        % @TITLE ~-- Import a far field from CST \code{.ffs} file
        % 
        % @DESCRIPTION
        %   Reads .ffs ("far field source") file generated in CST software and
        % returns the far field in the \hyperref[secFieldStructDescription]{Field structure}.
        % Also saves the read data (structures Field and Info) in a .mat file with
        % the same name as .ffs file and to the same directory.
        %   FFS files of version 3.0 are supported only, which supports field 
        % definition at multiple frequencies.
        % 
        % @SYNTAX
        % 
        % Field = TCSTInterface.ReadFarFieldSourceFile(FileName)
        % Field = TCSTInterface.ReadFarFieldSourceFile(Directory)
        % Field = TCSTInterface.ReadFarFieldSourceFile(Directory, FileMask)
        % Field = TCSTInterface.ReadFarFieldSourceFile(Directory, FileMask, NormalizationType)
        % Field = TCSTInterface.ReadFarFieldSourceFile(FileName, [], NormalizationType)
        % [Field, Info] = TCSTInterface.ReadFarFieldSourceFile(___)
        % 
        % @ARGUMENTS
        % 
        % @ARG FileName - File name of the .ffs file including the file extension
        %      and (if required) absolute or relative path. \refmetarg{FileName} can be
        %      a cell array where each cell contains a .ffs file name. In this case
        %      function will return array of \refmetret{Field} and \refmetret{Info}
        %      structures with elements corresponding to the files listed in
        %      \refmetarg{FileName}.
        % 
        % @ARG Directory - First argument of the function can also be a directory
        %      where several .ffs files are located. In this case the function will
        %      read all .ffs files found in this directory. An optional argument
        %      \refmetarg{FileMask} can be provided, which allows to filter .ffs files
        %      to be read.
        % 
        % @ARG FileMask - An optional argument which can be used to filter .ffs
        %      files in the directory \refmetarg{Directory}. E.g., with
        %      \code{FileMask='Co*.txt'} the function will read only ffs files
        %      starting with "Co" and which have an extension ".txt" (extension
        %      ".txt" is used sometimes instead of typical CST's ".ffs")
        % 
        % @ARG NormalizationType - A string specifying the far field normalization. Can be
        %      \code{'Gain'}, \code{'RealizedGain'}, \code{'Directivity'}, \code{'Unity'} or
        %      \code{'None'}. The accepted power, stimulated power, radiated power, the
        %      maximum amplitude of the far field, or unity are used to normalize the
        %      field for each \refmetarg{NormalizationType} correspondingly. The
        %      normalization is performed per frequency. The default value is
        %      \code{'None'}, so no normalization is applied.
        % 
        % @RETURN
        %
        % @RET Field - \hyperref[secFieldStructDescription]{Field structure} with
        %      the imported far field. Can be an array of structures if several .ffs
        %      files were read.
        % 
        % @RET Info - A structure with additional information extracted from the
        %      .ffs file. It contains info about frequency, radiated power,
        %      accepted power, stimulated power, as well as origin and axis of the
        %      coordinate system in which the far field was calculated.
        % 
        % @SEEALSO 
        %   \nameref{secFieldStructDescription}
        % 
        % @AUTHOR    Oleg Iupikov <oleg.iupikov@chalmers.se, lichne@gmail.com>
        % @VERSION   2.0
        % @FIRSTEDIT 11-2012
        % @LASTEDIT  19-10-2017

        %    11-2012: first version
        %    12-2012: + support multifrequency ffs v3.0
        % 07-03-2013:   improved performance ~4x
        % 19-10-2017: + normalization to gain or directivity

            % Info = [];
            if (nargin<2) || isempty(FileMask),  FileMask = '*.ffs'; end
            if nargin<3, NormalizationType = 'None'; end

            ValidStrings = {'Gain', 'RealizedGain', 'Directivity', 'Unity', 'None'};
            NormalizationType = validatestring(NormalizationType, ValidStrings, mfilename, 'NormalizationType', 3);

            if ~iscell(FileOrDir)
                % if FileOrDir is directory
                if TCSTInterface.isfolder(FileOrDir), 
                    fls = dir(fullfile(FileOrDir, FileMask));
                    Files = cell(length(fls)-2, 1);
                    for ifl=1:length(fls)
                        if strcmpi(fls(ifl).name,'.')||strcmpi(fls(ifl).name,'..'), continue; end
                        Files{ifl} = fullfile(FileOrDir, fls(ifl).name);
                    end
                else % if it is one '.ffs' file
                    [~, ~, ext] = fileparts(FileOrDir);
                    assert(strcmpi(ext,'.ffs')||strcmpi(ext,'.txt'), 'File must be ".ffs" or ".txt"');
                    Files = {FileOrDir};
                end    
            else
                Files = FileOrDir;
            end

            for ifl=1:length(Files)
                [pathstr, name, ext] = fileparts(Files{ifl});
                assert(strcmpi(ext,'.ffs')||strcmpi(ext,'.txt'), 'File must be ".ffs" or ".txt"');
                if nargout>0,   [Field(ifl), Info(ifl)] = localProcessFile(pathstr, [name ext], NormalizationType); %#ok<AGROW>
                else, localProcessFile(pathstr, [name ext], NormalizationType);
                end
            end

            function [Field, Info] = localProcessFile(Directory, FileName, NormalizationType,SilentMode)

                if nargin<4, SilentMode = false; end
                Info = [];

                [fid, message] = fopen(fullfile(Directory, FileName));
                if fid<0, error('Can''t open file. System message: ''%s''', message); end
                Info.FileName = [Directory filesep FileName];
                NTh = [];

                try
                    % ====================== READ DATA FROM FILE ==========================
                    % read one line from the file
                    tline = fgets(fid);
                    DataReading = false;
                    Ver = 0;
                    while ischar(tline)

                        % ++++++++++++ HEADERS READING ++++++++++++++ 
                        if ~DataReading,
                            % check version
                            if contains(tline, 'Version:')
                                Ver = fscanf(fid, '%f\n',1);
                                if ~((Ver==1.1)||(Ver==3.0)), error('Version %2.1f is not supported.',Ver); end
                            end

                            % ---------------------- ver 3.0 ------------------------------
                            if (Ver==3),
                                % CS position
                                if strcmpi(deblank(tline), '// Position')
                                    tmp = fscanf(fid, '%f\n',3);
                                    Info.CoorSysOrigin = tmp;
                                end
                                % CS Z axis
                                if strcmpi(deblank(tline), '// zAxis')
                                    tmp = fscanf(fid, '%f\n',3);
                                    Info.CoorSysZAxis = tmp;
                                end
                                % CS X axis
                                if strcmpi(deblank(tline), '// xAxis')
                                    tmp = fscanf(fid, '%f\n',3);
                                    Info.CoorSysXAxis = tmp;
                                end
                                % Number of frequencies
                                if strcmpi(deblank(tline), '// #Frequencies')
                                    NFreq = fscanf(fid, '%f\n',1);
                                    Info.Freq = nan(NFreq,1);
                                    Info.RadiatedPower = nan(NFreq,1);
                                    Info.AcceptedPower = nan(NFreq,1);
                                    Info.StimulatedPower = nan(NFreq,1);
                                    iFr = 1;
                                end

                                if strcmpi(deblank(tline), '// Radiated/Accepted/Stimulated Power , Frequency')
                                    for n=1:NFreq,
                                        Info.RadiatedPower(n) = fscanf(fid, '%f\n',1);
                                        Info.AcceptedPower(n) = fscanf(fid, '%f\n',1);
                                        Info.StimulatedPower(n) = fscanf(fid, '%f\n',1);
                                        Info.Freq(n) = fscanf(fid, '%f\n',1);
                                        Field.Freq(n) = Info.Freq(n);
                %                         tline = fgets(fid); % just read empty line
                                    end
                                end

                            % ------------------------ ver 1.1 ----------------------------
                            elseif (Ver==1.1),
                                error('Should be extended for multifrequency files...');
                %                 if ~isempty(strfind(tline, 'Frequency'))
                %                     Field.Freq = fscanf(fid, '%f\n',1);
                %                 end
                            else
                %                 error('Hm...');
                            end
                            % -------------------------------------------------------------

                            % number of samples
                            if contains(tline, '>> Total #phi samples, total #theta samples')
                                d = fscanf(fid, '%i %i\n',2);
                                NPh = d(1);  NTh = d(2);
                            end
                            % if data started
                            if contains(tline, '>> Phi, Theta, Re(E_Theta), Im(E_Theta), Re(E_Phi), Im(E_Phi)')
                                DataReading = true;
                                if ~SilentMode,
                                    fprintf('Reading "%s": Freq = %i of %i (%2.3f MHz)\n',FileName, iFr, NFreq, Info.Freq(iFr)/1e6);
                                end
                                continue; % if we read all samples in one block
                            end

                        % ++++++++++++ DATA READING ++++++++++++++    
                        else
                            if isempty(NTh), error('Number of Theta and Phi samples was not found.'); end

                            d = fscanf(fid, '%g %g %g %g %g %g', [6 NTh*NPh]);
                            if size(d,2)~=NTh*NPh, error('The file actually has less field samples then specified. Is the file corrupted?'); end
                            Theta=d(2,:); Phi=d(1,:); EthRe=d(3,:); EthIm=d(4,:); EphRe=d(5,:); EphIm=d(6,:);
                            Field.E(:,:,1,1,iFr) = zeros(NTh,NPh); % ro-component
                            Field.E(:,:,2,1,iFr) = reshape(EthRe+1i*EthIm, NTh, NPh); % theta component
                            Field.E(:,:,3,1,iFr) = reshape(EphRe+1i*EphIm, NTh, NPh); % phi component
                            iFr = iFr+1;
                            DataReading = false; % find next sub-header
                            % make sure that we will find number of samples
                            if iFr<=NFreq,  NTh=[];  NPh=[];  end
                        end 
                        % read next line from the file
                        tline = fgets(fid);
                    end
                    % =====================================================================

                    % ============= REARRANGE THE DATA TO FIELD STRUCTURE =================
                    if isempty(NTh), error('The file actually has less frequencies than specified in its header. Is the file corrupted?'); end
                    Field.THETA = reshape(Theta/180*pi, NTh, NPh);
                    Field.PHI = reshape(Phi/180*pi, NTh, NPh);

                    Field.VectorComponents = 'theta-phi';
                    Field.PortImpedance = []; % not defined
                    Field.NearFar = 'far';
                    Field.GridType = 'spherical';
                    Field.GridSymmetry = 'unsymmetrical';
					
                    if exist('TCoorSys','file')
                        Field.CoorSys = TCoorSys( Info.CoorSysXAxis.',[],Info.CoorSysZAxis.', Info.CoorSysOrigin.');
                    end
                    % =====================================================================

                    % ===================== NORMALIZE THE FIELD ===========================
                    if any(strcmpi(NormalizationType,{'Gain','RealizedGain'})) && any(abs(Info.AcceptedPower)<1000*eps),
                        warning('TCSTInterface:ReadFFS:NoAcceptedPower', '"NormalizationType" is set to ''%s'', however the antenna accepted power is 0. Therefore the field normalization was not performed.',NormalizationType);
                        NormalizationType = 'None';
                    end
                    for ifr=1:NFreq,
                        switch NormalizationType,
                            case 'Gain',
                                PNF = 60*Info.AcceptedPower(ifr); % PNF = Power Normalization Factor
                            case 'RealizedGain',
                                PNF = 60*Info.StimulatedPower(ifr);
                            case 'Directivity',
                                PNF = 60*Info.RadiatedPower(ifr);
                            case 'Unity',
                                PNF = max(max( sum(abs(Field.E(:,:,:,:,ifr)).^2, 3) )); 
                            case 'None',
                                PNF = 1;
                        end
                        Field.E(:,:,:,:,ifr) = Field.E(:,:,:,:,ifr) / sqrt(PNF);
                    end
                    % =====================================================================

                catch ME
                    fclose(fid);
                    rethrow(ME);
                end

                fclose(fid);
            end
            
        end
        
        function str = SecToString(sec, ShowMSec)
        % @DESCRIPTION
        %   Converts seconds to a string in convenient format.
        % 
        % @SYNTAX
        % str = TCSTInterface.SecToString(sec)
        % str = TCSTInterface.SecToString(sec, ShowMSec)
        % 
        % @EXAMPLES
        % TCSTInterface.SecToString(5581.77)        % '1 hours 33 min 02 sec'
        % TCSTInterface.SecToString(5581.77, true)  % '1 hours 33 min 01 sec 770 msec'
            
            if nargin<2,  ShowMSec = false;  end
            [Y, M, D, H, MN, S] = datevec(sec/3600/24);
            str = '';
            if Y>0, str = sprintf('%s%i years ',str,Y); end
            if M>0, str = sprintf('%s%i months ',str,M); end
            if D>0, str = sprintf('%s%i days ',str,D); end
            if H>0, str = sprintf('%s%i hours ',str,H); end
            if MN>0, str = sprintf('%s%02i min ',str,MN); end
            if ShowMSec,
                Sr = floor(S);
                if Sr>0,  str = sprintf('%s%02i sec ',str,Sr);  end
                str = sprintf('%s%03i msec ',str,round((S-Sr)*1000));
            else
                str = sprintf('%s%02i sec ',str,round(S));
            end
            str = str(1:end-1);
        end
        
        function str = FreqToString(Freq, ForFileName)
        % @DESCRIPTION
        %   Converts frequency in Hz to a string in convenient format.
        % 
        % @EXAMPLES
        % TCSTInterface.FreqToString(33)          % '33 Hz'
        % TCSTInterface.FreqToString(33.5e8)      % '3.35 GHz'
            
            if nargin<2, ForFileName = false; end
            
            if Freq<1e3, str = sprintf('%.6g Hz',Freq); 
            elseif Freq<1e6, str = sprintf('%.6g kHz',Freq/1e3); 
            elseif Freq<1e9, str = sprintf('%.6g MHz',Freq/1e6); 
            elseif Freq<1e12, str = sprintf('%.6g GHz',Freq/1e9); 
            elseif Freq<1e15, str = sprintf('%.6g THz',Freq/1e12); 
            else, str = sprintf('%e Hz',Freq);
            end
            
            if ForFileName,
                str = strrep(str, '.','p');
                str = strrep(str, ' ','');
            end
        end
        
        function str = FreqListToString(Freqs, MaxFreqsToShow)
        % @EXAMPLES
        % TCSTInterface.FreqListToString(1:3)             % '1 Hz, 2 Hz, 3 Hz'
        % TCSTInterface.FreqListToString([1:100]*1e8)     % '100 MHz, 200 MHz, 300 MHz, ..., 10 GHz'
        % TCSTInterface.FreqListToString([1:100]*1e8, 2)  % '100 MHz, ..., 10 GHz'
            
            if nargin<2, MaxFreqsToShow = 4; end
            NFreqs = length(Freqs);
            str = '';
            if isempty(Freqs), return; end
            if NFreqs<=MaxFreqsToShow,
                for k=1:NFreqs,
                    str = sprintf('%s%s, ', str, TCSTInterface.FreqToString(Freqs(k)));
                end
                str = str(1:end-2); % remove last coma
            else % MaxFreqsToShow+1 or more
                for k=1:MaxFreqsToShow-1,
                    str = sprintf('%s%s, ', str, TCSTInterface.FreqToString(Freqs(k)));
                end
                str = str(1:end-2); % remove last coma
                str = sprintf('%s, ..., %s', str, TCSTInterface.FreqToString(Freqs(end)));
            end
        end
        
        function str = VectorToString(Vec, MaxNumbersToShow)
        % @EXAMPLES
        % TCSTInterface.VectorToString(1:10)            % '1, 2, 3, ..., 10'
        % TCSTInterface.VectorToString([8:100]*1e8)     % '800e6, 900e6, 1e9, ..., 10e9'
        % TCSTInterface.VectorToString([8:100]*1e8, 2)  % '800e6, ..., 10e9'
            
            if nargin<2, MaxNumbersToShow = 4; end
            NElem = length(Vec);
            str = '';
            if isempty(Vec), return; end
            if NElem<=MaxNumbersToShow,
                for k=1:NElem,
                    str = sprintf('%s%s, ', str, locNS(Vec(k)));
                end
                str = str(1:end-2); % remove last coma
            else % MaxFreqsToShow+1 or more
                for k=1:MaxNumbersToShow-1,
                    str = sprintf('%s%s, ', str, locNS(Vec(k)));
                end
                str = str(1:end-2); % remove last coma
                str = sprintf('%s, ..., %s', str, locNS(Vec(end)));
            end
            
            function str = locNS(Val)
                Pow = floor(log10(Val)/3 + 100*eps);
                if Pow==0 || Pow==-1,
                    str = sprintf('%g', Val);
                else
                    str = sprintf('%ge%d', Val/1000^Pow, 3*Pow);
                end
            end
        end
        
        function [coef, unit] = GetFreqUnitCoefficient(StringWithUnits,RequireUnits)
        % @DESCRIPTION
        %   Gets units factor from the frequency string.
        % 
        % @SYNTAX
        % [coef, unit] = GetFreqUnitCoefficient(StringWithUnits, RequireUnits)
        %
        % @ARGUMENTS
        % @ARG StringWithUnits - String to to get frequency units from, e.g. \code{'1.2 GHz'}.
        % @ARG RequireUnits - If \code{true}, throws error if frequency units were not
        %      found. If \code{false}, the error will be ignored and units Hz assumed.
        %      Default value: \code{false}.
        % 
        % @EXAMPLES
        % % This returns:  coef = 1.0000e+09;  unit = 'GHz'
        % [coef, unit] = TCSTInterface.GetFreqUnitCoefficient('20 GHz') 
        % % This ignors wrong units and returns result in Hz:  coef = 1;  unit = 'Hz'
        % [coef, unit] = TCSTInterface.GetFreqUnitCoefficient('20 WHz') 
        % % This gives the error: 'Cannot decode frequency with units "20 WHz": unit was not found.'
        % [coef, unit] = TCSTInterface.GetFreqUnitCoefficient('20 WHz',true) 
            
            if nargin<2, RequireUnits = false; end
            % find units
            d = regexp(StringWithUnits,'((\s+Hz)|(kHz)|(MHz)|(GHz)|(THz)|())*','match');
            assert(length(d)<=1, 'Cannot decode frequency with units "%s": more than 1 unit found.',StringWithUnits);
            if isempty(d), 
                assert(~RequireUnits, 'Cannot decode frequency with units "%s": unit was not found.',StringWithUnits);
                unit = 'Hz'; 
            else
                unit = strtrim(d{1}); 
            end
            switch lower(unit),
                case 'hz',  coef = 1;
                case 'khz', coef = 1e3;
                case 'mhz', coef = 1e6;
                case 'ghz', coef = 1e9;
                case 'thz', coef = 1e12;
                otherwise, error('Cannot decode frequency with units "%s": not supported units.',StringWithUnits);
            end
        end
        
        function [val, coef, unit] = FreqStringToHertz(str,RequireUnits) 
            if nargin<2, RequireUnits = false; end
            % find number
            d = regexp(str,'(\d|\.|,|e|E|+|-)*','match'); 
            assert(~isempty(d), 'Cannot decode frequency with units "%s": number was not found.',str);
            assert(length(d)==1, 'Cannot decode frequency with units "%s": more than 1 number found.',str);
            num = str2double( strrep(d{1},',','.') );
            % find units
            [coef, unit] = TCSTInterface.GetFreqUnitCoefficient(str,RequireUnits);
            % value in meters
            val = num*coef;    
        end
        
        function numRunIDs = RunIDsStrToNum(strRunIDs)
            
            % check the argument
            ErrMsg = '"strRunIDs" must be a string or a cell array with strings, each of which is representing a CST RunID in format "3D:RunID:<n>".';
            assert(ischar(strRunIDs)||iscell(strRunIDs), ErrMsg);
            if ~iscell(strRunIDs),  strRunIDs = {strRunIDs};  
            else, assert(ischar(strRunIDs{1}), ErrMsg);
            end
            
            % parse the RunID string(s)
            d = regexp(strRunIDs, '3D:RunID:(\d+)', 'tokens');
            
            % check that all of them are of expected format
            WrongFormat = find( cellfun(@(c)isempty(c), d) );
            if ~isempty(WrongFormat),
                error([ErrMsg ' "' strRunIDs{WrongFormat(1)} '" doesn''t match the expected format.']);
            end
            
            % convert to numbers
            numRunIDs = cellfun(@(c)str2double(c{1}{1}), d);
        end
        
        % -----------------------------------------------------------------
        % Function for compatibility with Matlab versions below 9.3 (R2017b)
        % -----------------------------------------------------------------
        function Res = isfolder(DirName)
            if verLessThan('matlab','9.3')
                Res = isdir( TCSTInterface.GetFullPath(DirName) );
            else
                Res = isfolder(DirName);
            end
        end
        
%         function WriteTouchstoneFile(FileName, S, Freqs, ReferenceImpedanceOhm, strFreqUnits)
%             % WriteTouchstoneFile(FileName, S, Freqs, [ReferenceImpedanceOhm], [strFreqUnits])
%             %   S(NPorts,NPorts,NFreqs)
%             %   Freqs are in [Hz], if not specified otherwise 
%             %   ReferenceImpedanceOhm = 50 Ohm by default
%             
%             assert(ndims(S)<=3, '"S" has more than 3 dimensions which is not supported. The S-matrix format should be S(NPorts,NPorts,NFreqs).');
%             assert(size(S,3)==length(Freqs), 'Number of S-marices in "S" do not correspond to the number of frequencies in "Freqs".');
%             if nargin<4, ReferenceImpedanceOhm = 50; end
%             if nargin<5, strFreqUnits = 'Hz'; end
%             ValidStrings = {'Hz','kHz','MHz','GHz','THz'};
%             strFreqUnits = validatestring(strFreqUnits, ValidStrings, 'WriteTouchstoneFile', 'strFreqUnits', 5);
%             
%             
%             % write file
%             [fid, message] = fopen(FileName,'w');
%             if fid<0, error('Can''t open file for writing. System message: ''%s''', message); end
%             
%             fprintf(fid, '! TOUCHSTONE file generated by the CST interface for Matlab\n');
%             c = clock; 
%             fprintf(fid, '! Date and time: %i-%02i-%02i %02i:%02i:%02i\n', floor(c));
%             fprintf(fid, '# %s S RI R %i\n', strFreqUnits, ReferenceImpedanceOhm);
%             
%             fclose(fid);
%             
%             
%         end
            
    end
    
    % =====================================================================
    % ====================== PRIVATE methods ==============================
    % =====================================================================
    methods (Access = private)
%     methods % TEMP!

        function nbytes = fprintf(this, varargin)
            
            if this.FSilentMode,
                nbytes = 0;
                return;
            end
            
%             time = sprintf('[%s] ', datestr(now,'HH:MM:SS'));
%             if ischar(varargin{1}), % no file id
%                 varargin{1} = [time varargin{1}];
%             else % with a file id
%                 varargin{2} = [time varargin{2}];
%             end
            nbytes = fprintf(varargin{:});
        end
        
        function [TempDir, MacroFile, DataFile] = GetTempDirAndFiles(~,FileNamePart)
            TempDir   = fullfile(tempdir,  'CSTMatlabInterfaceTempFiles');
            if ~exist(TempDir, 'dir'),  mkdir(TempDir);  end
            if nargin>=2,
                MacroFile = fullfile(TempDir, [FileNamePart 'Wrapper.mcr']);
                DataFile  = fullfile(TempDir, [FileNamePart 'Data.txt']);
            end
        end
                
        function CheckIfValueSpecified(~, varargs, iarg, ParClasses, ParAttributes, FuncName)
            % CheckIfValueSpecified(varargs, iarg)
            % CheckIfValueSpecified(varargs, iarg, classes)
            % CheckIfValueSpecified(varargs, iarg, classes, attributes)
            % CheckIfValueSpecified(varargs, iarg, classes, attributes, funcName)
            % 
            % "classes,attributes,funcName" are as for "validateattributes".  
            % 
            % Examples:
            %   CheckIfValueSpecified(varargin, 2, 'PointIndex', {'integer'}, {'positive','finite','nonempty'})

            ParName = varargs{iarg};
            assert(length(varargs)>=iarg+1, 'A value is expected for the parameter "%s".', ParName);

            if nargin<4, return; end

            if nargin<5, ParAttributes = {'nonempty'}; end
            if nargin<6, FuncName = ''; end


            % if ParClasses provided, check ParName
            ParVal = varargs{iarg+1};
            validateattributes(ParVal,ParClasses,ParAttributes,FuncName,ParName,iarg);
        end

        function [Lines, S] = ReadTextFileContentToLines(~,FileName)
            fidR = fopen(FileName,'r');
            if fidR<0, error('CSTInterface:CantOpenFile', 'Can''t open file "%s". Does it exist?', FileName); end
            try
                S = fread(fidR, '*char');
            catch ME
                fclose(fidR);
                rethrow(ME);
            end
            fclose(fidR);
            assert(~isempty(S), 'CSTInterface:FileEmpty', 'The file "%s" is empty.', FileName);
            Lines = strsplit(S.',newline).'; % split on lines
        end
        
        % check if project is assigned to this.FProj
        function Res = CheckProjectIsOpen(this, ThrowErrorIfNot)
            if nargin<2, ThrowErrorIfNot = false; end
            Res = ~isempty(this.FProj);
            % even if this.FProj is not empty, the project could be closed, and any call this.FProj.invoke(...) will give error "The RPC server is unavailable". Check if it is a case  
            if Res,
                try
                    this.FProj.invoke('GetApplicationName'); 
                catch
                    Res = false;
                    if iscom(this.FProj),
                        this.FProj.release;
                    end
                    this.FProj = []; % mark that the project object is not availavle 
                end
            end
            assert(~ThrowErrorIfNot||Res, 'There is no open project.');
        end
        
        function [Res, Proj] = CheckFileNameOfOpenProject(this, FullFileName)
            FullFileName = TCSTInterface.GetFullPath(FullFileName);
            % get current active project
            Proj = this.FApp.Active3D;
            % if getting the project failed, return
            if ~iscom(Proj) && ~isinterface(Proj),
                Res = false;
                Proj = [];
                return
            end
            % check if corresponds to the requested one 
            ProjectPathAndName = Proj.invoke('GetProjectPath', 'Project');
            [OpennedProjPath,   OpennedProjName]   = fileparts(ProjectPathAndName);
            [RequestedProjPath, RequestedProjName] = fileparts(FullFileName);
            % if the active proj and the requested proj are same 
            Res = strcmpi(fullfile(OpennedProjPath,OpennedProjName), fullfile(RequestedProjPath,RequestedProjName));
        end
        
        function ObjectFullNames = CheckObjectFullNameAndExistance(this, ObjectFullNames)
            assert(ischar(ObjectFullNames)||iscell(ObjectFullNames), '"ObjectFullName" must be a string or cell array of strings.');
            if ~iscell(ObjectFullNames), ObjectFullNames = {ObjectFullNames}; end
            this.ObjectExist(ObjectFullNames, true);
        end
        
        % NumberOfTries times tries to calls method Command of the solver, returns if simulation succeded 
        function Res = RunSolver(this, NumberOfTries, Command)
            TryN = 1;
            while true,
                Res = this.FSolver.invoke(Command);
                if Res~=0,
                    break;
                else
                    TryN = TryN+1;
                    if TryN>NumberOfTries,
                        error('Could not solve in %i try(ies).', NumberOfTries);
                    end
                    fprintf('\n');
                    BTState = warning('backtrace');
                    warning('backtrace','off');
                    warning('Solver error. Try %i', TryN);
                    warning('backtrace',BTState.state);
                    pause(5);
                end
            end            
        end
        
        function setCalculateYZMatrices(this, State)
            if isempty(State), return; end
            if this.CheckProjectIsOpen,
                PostProcess1D = this.FProj.invoke('PostProcess1D');
                PostProcess1D.invoke('ActivateOperation', 'yz-matrices', State);
            end
        end
        
        function [strRunID, Flag] = GetRunIDString(~, RunID)
            if ischar(RunID),  
                Flag = (RunID(1)=='~') || (RunID(1)=='-');
                if Flag,
                    strRunID = RunID(2:end);
                else
                    strRunID = RunID;
                end
            elseif isnumeric(RunID),
                assert(isscalar(RunID), '"RunID" must be a scalar.');
                Flag = (RunID<0) || isnan(RunID);
                if isnan(RunID),  RunID = 0;  end % NaN is equivalent of -0, i.e. exclude RunID 0
                strRunID = sprintf('3D:RunID:%i', abs(RunID));
            else
                error('Invalid RunID.');
            end
        end
        
        function RunIDs = RunIDsToCellArray(~, RunIDs)
            if isnumeric(RunIDs),  RunIDs = num2cell(RunIDs);  end
            if ischar(RunIDs),     RunIDs = {RunIDs};          end
            % if it is already a cell array, check that each its element is a string or a scalar 
            if iscell(RunIDs),
                for n=1:length(RunIDs),
                    assert(isscalar(RunIDs{n}) || ischar(RunIDs{n}),  'When "RunIDs" is a cell array, its each element must be a string or a scalar number.')
                end
            end
        end
        
        function [strRunIDs, Flags] = GetRunIDsStrings(this, RunIDs)
            RunIDs = this.RunIDsToCellArray(RunIDs);
            NIDs = length(RunIDs);
            strRunIDs = cell(NIDs,1);
            Flags = false(NIDs,1);
            for n=1:NIDs,
                [strRunIDs{n}, Flags(n)] = this.GetRunIDString(RunIDs{n});
            end
        end
        
        function [IncludeRunIDs, ExcludeRunIDs] = GetRunIDsStringsIncludedExcluded(this, RunIDs)
            RunIDs = this.RunIDsToCellArray(RunIDs);
            NIDs = length(RunIDs);
            IncludeRunIDs = cell(NIDs,1);   NIncl = 0;
            ExcludeRunIDs = cell(NIDs,1);   NExcl = 0;
            for n=1:NIDs,
                [strRunID, Flag] = this.GetRunIDString(RunIDs{n});
                % Flag denotes if we want to exclude RunID
                if Flag,  NIncl = NIncl+1;   ExcludeRunIDs{NIncl} = strRunID; 
                else,     NExcl = NExcl+1;   IncludeRunIDs{NExcl} = strRunID; 
                end
            end
            IncludeRunIDs(cellfun(@isempty, IncludeRunIDs)) = [];
            ExcludeRunIDs(cellfun(@isempty, ExcludeRunIDs)) = [];
        end
        
        function RunIDs = GetFilteredRunIDs(this, RunIDs, FilterRunIDs, TreeItem)
            % RunIDs are expected to be in CST format: cell array of strings like "3D:RunID:0" 
            % FilterRunIDs can be a numeric array, string, or a cell array of RunIDs in CST format 
            % TreeItem is optional, to form more informative error text in case of an error 
            
            if nargin<4, TreeItem = []; end
            
            % Get filters (incuded and excluded RunIDs)
            [IncludeRunIDs, ExcludeRunIDs] = this.GetRunIDsStringsIncludedExcluded(FilterRunIDs);
            
            % ... If IncludeRunIDs are specified, keep only them 
            if ~isempty(IncludeRunIDs),
                inds = ismember(RunIDs, IncludeRunIDs);
                RunIDs = RunIDs(inds);
                if isempty(TreeItem),
                    assert(~isempty(RunIDs), 'There are no results left after filtring by specified "RunIDs". None of the specified RunIDs to include is present in the simulation.');
                else
                    assert(~isempty(RunIDs), 'There are no results left for the tree item "%s" after filtring by specified "RunIDs". None of the specified RunIDs to include is present in the simulation.', TreeItem);
                end
            end
            % ... If ExcludeRunIDs are specified, remove them
            if ~isempty(ExcludeRunIDs),
                inds = ismember(RunIDs, ExcludeRunIDs);
                RunIDs(inds) = [];
                if isempty(TreeItem),
                    assert(~isempty(RunIDs), 'There are no results left after filtring by specified "RunIDs". All available results have been excluded.');
                else
                    assert(~isempty(RunIDs), 'There are no results left for the tree item "%s" after filtring by specified "RunIDs". All available results have been excluded.', TreeItem);
                end
            end
        end
        
        function [QueriedX, iX] = FindQueriedX(this, X, QueryX, Verbose)
            
            % If we are interested in specific X only, select them 
            if ~isempty(QueryX),
                NX = length(QueryX); % overwrite number of X
                iX = zeros(NX,1);
                QueriedX = nan(NX,1); 
                for iqx=1:NX,
                    dx = abs(X-QueryX(iqx));
                    ind = find(dx<1e-6*QueryX(iqx), 1);
                    % if exact X is found, store it 
                    if ~isempty(ind),
                        QueriedX(iqx) = X(ind);
                    % ... otherwise, choose what to do based on the NearestOrInterpolated option 
                    else
                        % check boundaries
                        xmax = max(X);  xmin = min(X);
                        assert( (QueryX(iqx)<=xmax) && (QueryX(iqx)>=xmin), ...
                            'Value of requested QueryX=%g is out of bounds of the data X which is [%g %g].', QueryX(iqx), xmin, xmax);
                        % find index
                        switch this.NearestOrInterpolated
                            case 'Exact'
                                error('X=%g is not found in the simulation data. You may consider to set NearestOrInterpolated property to ''Neasrest'' or ''Interpolated''.', QueryX(iqx));
                            case 'Nearest'
                                ind = find(dx==min(dx), 1);
                                if Verbose
                                    BTState = warning('off','backtrace');
                                    warning('CSTInterface:NoExactX:Nearest', 'X=%g is not found in the simulation data. The closest point X=%g is taken instead.', QueryX(iqx), X(ind));
                                    warning(BTState.state,'backtrace');
                                end
                                QueriedX(iqx) = X(ind);
                            case 'Interpolated'
                                ind = nan;
                                QueriedX(iqx) = QueryX(iqx);
                            otherwise, 
                                error('Bug.');
                        end
                    
                    end
                    iX(iqx) = ind;
                end
                
                % if interpolation will be used for the data
                iPtsToInterpolate = isnan(iX);
                if any(iPtsToInterpolate) && Verbose,
                    this.fprintf('! An interpolation is used for the following data point(s): %s\n', ...
                        TCSTInterface.VectorToString(QueriedX(iPtsToInterpolate)));
                end
            else
                NX = length(X);
                iX = 1:NX;
                QueriedX = X;
            end
            
        end
        
        function [QueriedX, ResIDs, Info] = Get1DResultXFor1stRunID(this, TreeItem, QueryX)
            if nargin<3, QueryX = []; end
            QueriedX = [];  Info = [];
            
            % Get ResIDs from the TreeItem
            ResIDs = this.GetResultIDsFromTreeItem(TreeItem);
            if isempty(ResIDs), return; end
            
            % Get ResultTree object from the project and Get 1D result object for the 1st RunID  
            ResultTree = this.FProj.invoke('Resulttree');
            objRes = ResultTree.invoke('GetResultFromTreeItem', TreeItem, ResIDs{1});
            % Make sure it is 1D result
            ResType = objRes.invoke('GetResultObjectType');
            assert(~isempty(intersect(ResType,{'1D','1DC'})), 'Tree item "%s" is expected to have 1D result. However, it contains unsupported result type "%s".',TreeItem,ResType);
            Info.ResType = ResType;
            
            % Get all X
            X = objRes.invoke('GetArray','x');
            if isrow(X), X = X.'; end

            % If we are interested in specific X only, select them 
            QueriedX = this.FindQueriedX(X, QueryX, true);
            
            % get additional info
            Info.XLabel = objRes.invoke('GetXLabel');
            Info.YLabel = objRes.invoke('GetYLabel');
            Info.Title  = objRes.invoke('GetTitle');
            
            % For some reason TreeItemHasImpedance gives sometimes error if TreeItem doesn't have any impedance attached... 
            % What is the purpose of this function then? Anyway, we just wrap it in try...catch block ...
            try
                Info.TreeItemHasImpedance = ResultTree.invoke('TreeItemHasImpedance', TreeItem, ResIDs{1});
            catch
                Info.TreeItemHasImpedance = false;
            end
            
        end
        
        function CheckXIsSame(~, X1st, X, RunID, QueryX)
            
            % We need to check this only when QueryX is NOT specified. If
            % QueryX IS specified, X for all RunIDs ARE same.
            if ~isempty(QueryX), return; end
            
            % Check that result for this RunID has same X vector  
            assert( (numel(X)==numel(X1st)) && all((abs(X(:)-X1st(:))<1e-6*min(X1st(:)))), ...
                ['X for Run ID "%s" is not the same as for the previous ones. This is not supported.\n'...
                'You may consider specifying QueryX as an additional input argument together with setting NearesOrInterpolated to ''Interpolated'''], ...
                RunID ...
            )
        
        end
        
        function Yout = GetIndexedY(~, X, Y, iX, QueryX)
            
            % Interpolate data if it is required. Data points to interpolate are marked with NaN in iX. 
            iQueryXToInterpolate = isnan(iX);
            Yout = nan(length(iX),1);
            if ~isempty(iQueryXToInterpolate), % can be true only if QueryX is specified and NearesOrInterpolated='Interpolated' 
                Yout(iQueryXToInterpolate) = interp1(X,Y, QueryX(iQueryXToInterpolate));
                ind = ~isnan(iX);
            else
                ind = 1:NX;
            end
            Yout(ind) = Y(iX(ind));
            
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [D, Freqs, Info] = GetSYZParams(this, TreePath, varargin)
            % [D, Freq, Dref, Info] = GetSParams(this, TreePath, RunIDs, iRows, iCols, PortMode, FreqsToGet)
            
            this.CheckProjectIsOpen(true);
            if nargin>=3,  RunIDFilter = varargin{1};  else,  RunIDFilter = [];   end
            if nargin>=4,  iRows       = varargin{2};  else,  iRows = [];         end
            if nargin>=5,  iCols       = varargin{3};  else,  iCols = [];         end
            if nargin>=6,  PortMode    = varargin{4};  else,  PortMode = [];      end
            if nargin>=7,  FreqsToGet  = varargin{5};  else,  FreqsToGet = [];    end
            
            % Get ResultTree object from the project
            ResultTree = this.FProj.invoke('Resulttree');
            assert(ResultTree.invoke('DoesTreeItemExist',TreePath), 'There is no result path "%s".', TreePath);
            
            % Get list of all available X-params
            TreeItems = this.GetResultTreeItemChildren(TreePath);
            assert(~isempty(TreeItems), 'There are no results under path "%s".', TreePath);
            % EXCLUDE POSSIBLE FLOQUET PORTS (TODO: save X-matrix containing them to separate matrix)   
            iTreeItemsWithFloquetPort = contains(TreeItems, 'Zmin') | contains(TreeItems, 'Zmax');
            TreeItems(iTreeItemsWithFloquetPort) = [];
            % 
            NTreeItems = length(TreeItems);
            
            % Check that ports have several modes. The result format will depend on it.  
            PortsHaveSeveralModes = localCheckIfPortsHaveSeveralModes(TreeItems{1});
            % indicates if the resulting array will have modes as additional dimensions  
            SaveWithModes = PortsHaveSeveralModes && isempty(PortMode);
            
            % Find maximum marix indices
            MaxM = 1;  MaxN = 1;  MaxMMode = 1;  MaxNMode = 1;
            for iti = 1:NTreeItems,
                [m,mmode,n,nmode] = localGetElemIndices(TreeItems{iti});
                MaxM = max(MaxM,m);
                MaxN = max(MaxN,n);
                if SaveWithModes,
                    MaxMMode = max(MaxMMode,mmode);
                    MaxNMode = max(MaxNMode,nmode);
                end
            end
            
            % 
            if ~isempty(iRows),  MaxN = length(iRows(iRows<=MaxN));  end  
            if ~isempty(iCols),  MaxM = length(iCols(iCols<=MaxM));  end
            
            % Get number of results (runs for e.g. parametric sweep). 
            % Here we ASSUME that all tree items have same number of results. 
            RunIDs = this.GetResultIDsFromTreeItem(TreeItems{1});
            NResultsItem1 = length(RunIDs);
            % Filter ResIDs keepeng only that which user specified 
            RunIDs = this.GetFilteredRunIDs(RunIDs, RunIDFilter);
            NResults = length(RunIDs);
            
            % Get frequencies. 
            % Here we ASSUME that all tree items have same number of frequencies. 
            % This assumption will be checked for each TreeItem and RunID.
%             [Freqs, RunIDs, Info] = this.Get1DResultXFor1stRunID(TreeItems{1}, FreqsToGet); 
            %
            [~, ~, Info] = this.Get1DResultXFor1stRunID(TreeItems{1}, FreqsToGet);  
            objRes = ResultTree.invoke('GetResultFromTreeItem', TreeItems{1}, RunIDs{1});
            X = objRes.invoke('GetArray','x');    if isrow(X), X = X.'; end
            Freqs = this.FindQueriedX(X, FreqsToGet, true);
            %
            FreqsInPlotUnits = Freqs;
            % try to convert to Hz
            try
                coef = TCSTInterface.GetFreqUnitCoefficient(Info.XLabel); 
                Freqs = Freqs*coef;                                         
            catch ME,
                warning('TCSTInterface:CannotConvertFreqUnits', ...
                    ['Could not covert the frequecy units to Hz from the CST plot X-label ("%s"). The units displayed below may be wrong.\n', ...
                    'Error message:\n%s'], Info.XLabel, ME.message);
            end
            NFreq = length(Freqs);
            this.fprintf('Data for %i frequencies will be read:  %s\n', NFreq, TCSTInterface.FreqListToString(Freqs));
            
%             % Get number of results (runs for e.g. parametric sweep). 
%             % Here we ASSUME that all tree items have same number of results. 
%             NResultsItem1 = length(RunIDs);
%             % Filter ResIDs keepeng only that which user specified 
%             RunIDs = this.GetFilteredRunIDs(RunIDs, RunIDFilter);
%             NResults = length(RunIDs);
            
            % Init output X-matrix and TreeImpedance (= reference impedance for S-params)
            % We don't know for sure how many ports and modes we have, therefore initialize them with 1; the array will be increased by Matlab in loop... 
            % The data array D will be
            %  - 6D-array: D(m,mmode,n,nmode,iFreq,iResID), if port(s) have several modes AND PortMode is not specified 
            %  - 4D-array: D(m,n,iFreq,iResID), otherwise 
            if SaveWithModes,
                D = nan(MaxM, MaxMMode, MaxN, MaxNMode, NFreq, NResults);
                Info.DataIndexing = 'D(m,mmode,n,nmode,iFreq,iResID)';
                % we always initialize TreeItemImpedance, and if no TreeItem has it, we will assign [] to it  
                Info.TreeItemImpedance = nan(MaxM, MaxMMode, MaxN, MaxNMode, NFreq, NResults); 
                Info.TreeItemImpedanceIndexing = 'TreeItemImpedance(m,mmode,n,nmode,iFreq,iResID)';
            else
                D = nan(MaxM, MaxN, NFreq, NResults);
                Info.DataIndexing = 'D(m,n,iFreq,iResID)';
                % we always initialize TreeItemImpedance, and if no TreeItem has it, we will assign [] to it  
                Info.TreeItemImpedance = nan(MaxM, MaxN, NFreq, NResults);
                Info.TreeItemImpedanceIndexing = 'TreeItemImpedance(m,n,iFreq,iResID)';
            end
            
            % Process each tree item 
            for iti = 1:NTreeItems,
                TreeItem = TreeItems{iti};
                
                % Get matrix indices
                [m,mmode,n,nmode] = localGetElemIndices(TreeItem);
                
                % if we want skip some rows and columns
                if ~isempty(iRows) && ~any(iRows==n),  continue;  end  
                if ~isempty(iCols) && ~any(iCols==m),  continue;  end
                
                % If we are interested in one port mode only, don't read the data for unwanted modes  
                if PortsHaveSeveralModes && ~isempty(PortMode),
                    if mmode~=PortMode || nmode~=PortMode,  continue;  end
                end
                
                % Check that current tree item has same number of RunIDs
                ResIDsTest = this.GetResultIDsFromTreeItem(TreeItem);
                assert(~isempty(ResIDsTest), 'There are no results for the tree item "%s".', TreeItem);
                assert(length(ResIDsTest)==NResultsItem1, 'The tree item "%s" has %i results, which diffrent from the number of results for the 1st tree item "%s" (%i). This breaks the assumption that all tree items have same number of results, and the code must be revised!', ...
                    TreeItem, length(ResIDsTest), TreeItems{1}, NTreeItems);
                
                % Read data for each result
                for iID = 1:NResults,
                    RunID = RunIDs{iID};
                    this.fprintf('Reading "%s", result id "%s"...\n', TreeItem, RunID);
                    
                    % Get X-param object
                    objXPar = ResultTree.invoke('GetResultFromTreeItem', TreeItem, RunID);
                    % make sure that it is Result1DComplex Object 
                    assert(strcmp(objXPar.invoke('GetResultObjectType'), '1DC'), ...
                        'The result "%s" is not 1D complex result, as was expected.', TreeItem);
                    
                    % Get freq for this TreeItem and RunID
                    Frq = objXPar.invoke('GetArray','x');
                    if isrow(Frq), Frq = Frq.'; end
                    % Check that result for this RunID has same freq vector
                    this.CheckXIsSame(FreqsInPlotUnits, Frq, RunID, FreqsToGet);
                    
                    % Get Xmn data
                    [~, iFreqs] = this.FindQueriedX(Frq, FreqsToGet, false);
                    Dre = this.GetIndexedY(Frq, objXPar.invoke('GetArray','yre'), iFreqs, FreqsToGet);
                    Dim = this.GetIndexedY(Frq, objXPar.invoke('GetArray','yim'), iFreqs, FreqsToGet);
                    if SaveWithModes,
                        D(m,mmode,n,nmode,:,iID) = Dre + 1i*Dim; 
                    else
                        D(m,n,:,iID) = Dre + 1i*Dim;
                    end
                    
                    % As for CST2019, calling "TreeItemHasImpedance" gives an error "Error
                    % reading reference impedances from..." while trying to read Z-params.
                    % Sure, Z-params don't have reference impedance, but this is exactly what
                    % "TreeItemHasImpedance" should check! Anyway, we just wrap this
                    % block in try-catch as a work-around...
                    try
                        if nargout>=3 && ResultTree.invoke('TreeItemHasImpedance', TreeItem, RunID),
                            objTreeItemImpedance = ResultTree.invoke('GetImpedanceResultFromTreeItem', TreeItem, RunID);
                            Zre = this.GetIndexedY(Frq, objTreeItemImpedance.invoke('GetArray','yre'), iFreqs, FreqsToGet);
                            Zim = this.GetIndexedY(Frq, objTreeItemImpedance.invoke('GetArray','yim'), iFreqs, FreqsToGet);
                            if SaveWithModes,
                                Info.TreeItemImpedance(m,mmode,n,nmode,:,iID) = Zre + 1i*Zim;
                            else
                                Info.TreeItemImpedance(m,n,:,iID) = Zre + 1i*Zim;
                            end 
                        end
                    catch
                    end
                    
                end
                
            end
            
            % if all TreeItems have NO impedance
            if all(isnan(Info.TreeItemImpedance(:))),
                Info.TreeItemImpedance = [];
                Info = rmfield(Info, 'TreeItemImpedanceIndexing');
            % if all TreeItems impedances are same (e.g. lumped ports were used) 
            elseif all(abs(Info.TreeItemImpedance(~isnan(Info.TreeItemImpedance(:)))-Info.TreeItemImpedance(1))<eps(1000)),
                Info.TreeItemImpedance = Info.TreeItemImpedance(1);
                Info = rmfield(Info, 'TreeItemImpedanceIndexing');
            end
            
            Info.ResultIDs = RunIDs;
            if ~isempty(PortMode),
                Info.ResultsAreForPortMode = PortMode;
            else
                Info.ResultsAreForPortMode = 'All';
            end

            
            % ------- Helping functions -------
            function [Res, ItemName] = localCheckIfPortsHaveSeveralModes(TreeItem)
                % get element name without path
                Res = strsplit(TreeItem,'\');
                ItemName = Res{end};
                assert(~isempty(ItemName), 'Error while getting the tree item name from "%s"', TreeItem);
                Res = contains(ItemName,'(');
            end
            
            function [m,mmode,n,nmode] = localGetElemIndices(TreeItem)
                [SeveralModes, ItemName] = localCheckIfPortsHaveSeveralModes(TreeItem);
                % if it contains parenthesis, there are several modes present; use special regexp pattern for that  
                if SeveralModes,
                    d = regexp(ItemName, '\D+(\d+)\((\d+)\)\D+(\d+)\((\d+)\)', 'tokens');
                    assert(~isempty(d), 'Error while getting matrix indices from the tree item "%s".', ItemName);
                    assert(length(d{1})==4, 'Error while getting matrix indices from the tree item "%s".', ItemName);
                    m     = str2double(d{1}{1});
                    mmode = str2double(d{1}{2});
                    n     = str2double(d{1}{3});
                    nmode = str2double(d{1}{4});
                else
                    mmode = [];  nmode = []; % there are no port modes
                    d = regexp(ItemName, '\D+(\d+)\D+(\d+)', 'tokens');
                    assert(~isempty(d), 'Error while getting matrix indices from the tree item "%s".', ItemName);
                    assert(length(d{1})==2, 'Error while getting matrix indices from the tree item "%s".', ItemName);
                    m     = str2double(d{1}{1});
                    n     = str2double(d{1}{2});
                end
            end
            
        end
        
        function WriteNamedScriptToFile(~, ScriptName, ScriptFile, ScriptParams)
            % Read all content of this m-file
            FileName = [mfilename '.m'];
            fid = fopen(FileName, 'r');
            assert(fid>=0, 'Cannot open file "%s".', FileName);
            try
                S = fread(fid, '*char');
            catch ME
                fclose(fid);
                rethrow(ME);
            end
            fclose(fid);
            assert(~isempty(S), 'CSTInterface:TxtEmpty', 'The text read from "%s" is empty. Bug?', FileName);
            Lines = strsplit(S.',newline); % split on lines
            NLines = length(Lines);
            
            % find line with script name
            iLnSt = find(contains(Lines,sprintf('%%>>> Name: %s',ScriptName)));
            assert(~isempty(iLnSt),  'Internal error: Script with name "%s" was not found.',ScriptName);
            assert(length(iLnSt)==1, 'Internal error: Several scripts with name "%s" were found.',ScriptName);
            
            % make sure the directory for the script exists
            FilePath = fileparts(ScriptFile);
            if ~TCSTInterface.isfolder(FilePath), mkdir(FilePath); end
            
            % Write script content to the file ScriptFile
            fid = fopen(ScriptFile, 'w');
            assert(fid>=0, 'Cannot open file "%s" for writing.', ScriptFile);
            try
                iLn = iLnSt+1;
                while iLn <= NLines,
                    Line = Lines{iLn};  iLn = iLn+1;
                    if length(Line)>=4 && strcmp(Line(1:4),'%<<<'), % end of script found 
                        break; 
                    end
                    if length(Line)<2, continue; end
                    
                    % check if we have a parameter in this line
                    [pn, st, en] = regexp(Line, '\${((?:\w|\.|\(|\)|:)+)}', 'tokens','start','end');
                    if ~isempty(pn), % line contains parameter(s) with name(s) pn 
                        for ip=1:length(pn),
                            ParName = pn{ip}{1};
                            % find this param in specified list ScriptParams  
                            ip1 = find(strcmp(ScriptParams(:,1),ParName),1);
                            assert(~isempty(ip1), 'Internal error: Parameter "%s" was found in VB script, but its value is not specified.', ParName);
                            % replace the parameter with its value
                            Line = [Line(1:st-1) ScriptParams{ip1,2} Line(en+1:end)];
                        end
                    end
                    
                    % write the line to the script file
                    fprintf(fid,'%s',Line(2:end));
                end
            catch ME
                fclose(fid);
                rethrow(ME);
            end
            fclose(fid);
        end
    end
    
    % =====================================================================
    % ======================= PUBLIC methods ==============================
    % =====================================================================
    methods
        % -----------------------------------------------------------------
        % CONSTRUCTOR
        % -----------------------------------------------------------------
        function this = TCSTInterface(ProjectFile)
        % @DESCRIPTION 
        %   Class constructor. If the input argument (ProjectFile) is specified, the
        % project ProjectFile will be opened.
            
            if nargin>=1,
                if isempty(ProjectFile)
                    this.OpenProject();
                else
                    this.OpenProject(ProjectFile);
                end
            end
        
        end
        
        % -----------------------------------------------------------------
        % If CST application is already open, get its OLE object; 
        % otherwise, open CST
        % -----------------------------------------------------------------
        function ConnectToCSTOrStartIt(this)
        % @DESCRIPTION
        %   Creates CST actxserver or connects to the existing one.
        % 
        % @SYNTAX
        % Obj.ConnectToCSTOrStartIt();
            
            % try to get already opened CSR
            try
                this.FApp = actxGetRunningServer('CSTStudio.Application');
            catch % we are here if there is no running CST
                this.FApp = actxserver('CSTStudio.Application');
            end
        end
        
        % -----------------------------------------------------------------
        % Opens specified .cst file, or connects to the currently active project  
        % -----------------------------------------------------------------
        function OpenProject(this, FullFileName)
        % @DESCRIPTION
        %   Opens specified \code{.cst}-file, or connects to the currently active project
        % if no \refmetarg{FileName} specified. If the specified file is alrady open in the CST
        % GUI, it will be activated. After successfully calling this method, the
        % propery \refprop{Project} is set and can be used for manual project
        % manipulation.
        % 
        % @SYNTAX
        % Obj.ConnectToCSTOrStartIt();
        % Obj.ConnectToCSTOrStartIt([]);
        % Obj.ConnectToCSTOrStartIt(FileName);
        % 
        % @ARGUMENTS
        % @ARG FileName - a \code{.cst}-file to open. Can contain absolute or relative path. 
            
            if nargin<2,  FullFileName = [];  end
            
            % ensure the FileName is full (containing the full path) 
            if ~isempty(FullFileName),
                FullFileName = TCSTInterface.GetFullPath(FullFileName);
            end
            
            % Assign this.FApp if CST is opened, or start it first if not  
            if isempty(this.FApp),
                this.ConnectToCSTOrStartIt();
            end
            
            if isempty(FullFileName),
                this.FProj = this.FApp.Active3D;
                % if there is no currently open project, CST returns some strange 
                % long string instead of COM object
                if ~iscom(this.FProj) && ~isinterface(this.FProj),
                    this.FProj = [];
                    error('There is no active projects is CST. You can open .cst file by passing the file name as the argument to OpenProject.');
                end
            else
                % try to open. If it is already opened in another tab, it will be activated  
                this.FProj = this.FApp.OpenFile(FullFileName);
                % if not successful, try to see if the specified project is already opened
                if ~iscom(this.FProj) && ~isinterface(this.FProj),
                    [Res, Proj] = this.CheckFileNameOfOpenProject(FullFileName);
                    if Res,
                        this.FProj = Proj;
                    else
                        this.FProj = [];
                        error('There was an error upon opening "%s".', FullFileName);
                    end
                end  
                
            end  
            
            this.fprintf('Connected to project "%s.cst"\n', this.FProj.invoke('GetProjectPath','Project'));
        end
        
        % -----------------------------------------------------------------
        % Closes the project specifield by its file name or currently active project 
        % -----------------------------------------------------------------
        function CloseProject(this, FullFileName, SaveProject)
        % @DESCRIPTION
        %   Closes the project specifield by its file name \refmetarg{FileName}.
        % If no \refmetarg{FileName} is provided or it is empty, the method will 
        % attempt to close the currently open project. 
        % If \refmetarg{SaveProject} is set to true, the project will be saved first. 
        %   Note that if \refmetarg{SaveProject} is true and \refmetarg{FileName} is specified but is not  
        % currenctly opened, it will be openned first (this is because there
        % is no "activate project" functionality in CST VBA (as for CST2017) and \refmethod{OpenProject} should be
        % used to activate the project).
        % 
        % @SYNTAX
        % Obj.CloseProject();
        % Obj.CloseProject(FileName);
        % Obj.CloseProject(FileName, SaveProject);
        % 
        % @ARGUMENTS
        % @ARG FileName - the \code{.cst}-file of the open project to be closed. Can
        %      contain absolute or relative path. If not specified or empty - the currently
        %      active project will be closed.
        % @ARG SaveProject - (optional, default value = \code{false}) if set to \code{true}, the project will be saved first.
            
            if nargin<2,  FullFileName = [];  end
            if nargin<3,  SaveProject = false;  end % if false, then "Ask in CST if save is required" 
            
            % Assign this.FApp if CST is opened, or start it first if not
            if isempty(this.FApp),
                this.ConnectToCSTOrStartIt();
            end
            
            % if FullFileName is provided, ensure it is full (containing the full path) 
            if ~isempty(FullFileName),
                FullFileName = TCSTInterface.GetFullPath(FullFileName);
            % if FullFileName is NOT provided, determine currently open project to be closed     
            else
                % some project is expected to be open
                this.CheckProjectIsOpen(true);
                FullFileName = [this.FProj.invoke('GetProjectPath','Project') '.cst'];
            end
            
            if SaveProject,
                % if the currently open project is not FullFileName, open it, it will become active  
                [Res, this.FProj] = this.CheckFileNameOfOpenProject(FullFileName);
                if ~Res,
                    this.OpenProject(FullFileName);
                end
                this.FProj.invoke('Save');
            end
            
            % Close the project
            this.FApp.CloseProject(FullFileName);
            
        end
        
        % -----------------------------------------------------------------
        % Changes parameter in the project. The parameter must exist.
        % The structure will be updeted, if UpdateStructure=true or omitted. 
        % To prevent the structure to be updated (e.g. if multiple
        % parameters should be changed), set UpdateStructure to false. 
        % -----------------------------------------------------------------
        function ChangeParameter(this, ParName, ParValue, UpdateStructure)
        % @DESCRIPTION
        %   Changes a parameter with name \refmetarg{ParName} in the project.
        % The structure will be updeted, if \refmetarg{UpdateStructure}\code{=true} or omitted. 
        % To prevent the structure to be updated (e.g. if multiple
        % parameters should be changed), set \refmetarg{UpdateStructure} to \code{false}. 
        %   The parameter must exist. If a new parameter should be created, use
        % \refmethod{StoreParameter} instead.
        % 
        % @SYNTAX
        % Obj.ChangeParameter(ParName, ParValue);
        % Obj.ChangeParameter(ParName, ParValue, UpdateStructure);
        % 
        % @ARGUMENTS
        % @ARG ParName - Parameter name.
        % @ARG ParValue - Parameter value. Can be a numeric value or a string with any
        %      valid parametric expression.
        % @ARG UpdateStructure - (optional, default value = \code{true}) If set to
        %      \code{true}, the EM-model will be updated by processing the project history list.
            
            this.CheckProjectIsOpen(true);
            assert(ischar(ParName), '"ParName" must be a string.');
            
            % if UpdateStructure is omitted, the structure will be updated  
            if nargin<4, UpdateStructure = true; end
            
            % make sure that ParName exists in the project
            this.ParameterExists(ParName, true);
            % change parameter ParName
            this.StoreParameter(ParName, ParValue, UpdateStructure);
        end
        
        % -----------------------------------------------------------------
        % Same, but ParName and ParValue can be vectors now.
        % If UpdateStructure=true, the update is performes after last parameter change.
        % -----------------------------------------------------------------
        function ChangeParameters(this, ParNames, ParValues, UpdateStructure)
            this.CheckProjectIsOpen(true);
            if ischar(ParNames), ParNames = {ParNames}; end
            if nargin<4, UpdateStructure = true; end
            
            NPars = length(ParNames);
            for ipar=1:NPars,
                if ipar==NPars, Updt = UpdateStructure; else, Updt = false; end
                this.ChangeParameter(ParNames{ipar}, ParValues(ipar), Updt);
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function StoreParameter(this, ParName, ParValue, UpdateStructure)
        % @DESCRIPTION
        %   Same, as \refmethod{ChangeParameter}, except it doesn't require the parameter
        % to exist. If the paramter with name \refmetarg[\CurrentClass][ChangeParameter]{ParName}
        % doesn't exist, it will be created.
            
            this.CheckProjectIsOpen(true);
            assert(ischar(ParName), '"ParName" must be a string.');
            
            % if UpdateStructure is omitted, the structure will be updated  
            if nargin<4, UpdateStructure = true; end
            
            if iscell(ParValue)
                assert(isscalar(ParValue), 'When "ParValue" is a cell, it must be a scalar containing the parameter expression (string) or the parameter value (number).');
                this.StoreParameter(ParName, ParValue{1}, UpdateStructure);
                return
            end
            
            % store parameter ParName
            if ischar(ParValue)
                this.FProj.invoke('StoreParameter',ParName,ParValue); 
            elseif isnumeric(ParValue)
                this.FProj.invoke('StoreDoubleParameter',ParName,ParValue);
            else
                error('Wrong type of "ParValue".');
            end
            
            % update structure if required
            if UpdateStructure,
                Res = this.FProj.invoke('RebuildOnParametricChange',false,false);
                assert(Res, 'Error during the structure update on parametric change.');
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function StoreParameterWithDescription(this, ParName, ParValue, ParDescription, UpdateStructure)
            this.CheckProjectIsOpen(true);
            assert(ischar(ParName), '"ParName" must be a string.');
            assert(ischar(ParDescription), '"ParDescription" must be a string.');
            
            % if UpdateStructure is omitted, the structure will be updated  
            if nargin<5, UpdateStructure = true; end
            
            % store parameter ParName
            if isnumeric(ParValue),
                ParValue = num2str(ParValue);
            else
                assert(ischar(ParValue), '"ParValue" must be numeric or string.');
            end
            this.FProj.invoke('StoreParameterWithDescription',ParName,ParValue,ParDescription);
            
            % update structure if required
            if UpdateStructure,
                Res = this.FProj.invoke('RebuildOnParametricChange',false,false);
                assert(Res, 'Error during the structure update on parametric change.');
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function NParams = GetNumberOfParameters(this)
            this.CheckProjectIsOpen(true);
            NParams = this.FProj.invoke('GetNumberOfParameters');
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function Res = ParameterExists(this, ParamName, ThrowErrorIfNot)
            this.CheckProjectIsOpen(true);
            if nargin<3, ThrowErrorIfNot = false; end
            assert(ischar(ParamName), '"ParamName" must be a string.');
            
            Res = this.FProj.invoke('DoesParameterExist', ParamName);
            assert(~ThrowErrorIfNot||Res, 'Parameter "%s" doesn''t exist in the project.', ParamName);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function ParamNames = GetParameterList(this)
            this.CheckProjectIsOpen(true);
            
            NParams = this.GetNumberOfParameters();
            ParamNames = cell(NParams,1);
            for ipar=1:NParams,
                ParamNames{ipar} = this.FProj.invoke('GetParameterName',ipar-1);
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function ParamName = GetParameterNameByIndex(this, ParamIndex)
            this.CheckProjectIsOpen(true);
            assert(isscalar(ParamIndex), '"ParamIndex" must be a scalar number >= 1.');
            assert(ParamIndex>=1, '"ParamIndex" must be a scalar number >= 1.');
            
            ParamName = this.FProj.invoke('GetParameterName',ParamIndex-1);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [ParamIndex, AllParamNames] = GetParameterIndexByName(this, ParamName)
            this.CheckProjectIsOpen(true);
            assert(ischar(ParamName), '"ParamName" must be a string.');
            this.ParameterExists(ParamName, true);
            
            AllParamNames = this.GetParameterList();
            ParamIndex = find(strcmpi(AllParamNames,ParamName),1);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function ParamValue = GetParameterValue(this, ParamNameOrIndex)
            this.CheckProjectIsOpen(true);
            
            % Get parameter index
            if ischar(ParamNameOrIndex),  % if name is provided
                ParamIndex = this.GetParameterIndexByName(ParamNameOrIndex);
            else % if it is already index
                ParamIndex = ParamNameOrIndex;
                assert(ParamIndex>=1, 'Parameter index (%i) must be greater or equal to 1.', ParamIndex);
                NParams = this.GetNumberOfParameters();
                assert(ParamIndex<=NParams, 'Parameter index (%i) exceeds number of parameters in the project (%i)', ParamIndex, NParams);
            end
            
            ParamValue = this.FProj.invoke('GetParameterNValue', ParamIndex-1);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [ParamNames, ParamValues] = GetCurrentValueOfAllParameters(this)
            this.CheckProjectIsOpen(true);
            
            % Get parameters name
            ParamNames = this.GetParameterList();
            NParams = length(ParamNames);
            % Get parameters value
            ParamValues = nan(NParams,1);
            for ipar=1:NParams
                ParamValues(ipar) = this.GetParameterValue(ipar);
            end
            
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [ParamNames, ParamExpressions] = GetCurrentExpressionOfAllParameters(this)
            this.CheckProjectIsOpen(true);
            
            % Get parameters name
            ParamNames = this.GetParameterList();
            NParams = length(ParamNames);
            % Get parameters value
            ParamExpressions = cell(NParams,1);
            for ipar=1:NParams
                ParamExpressions{ipar} = this.GetParameterExpression(ipar);
            end
            
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function ParamExp = GetParameterExpression(this, ParamNameOrIndex)
            this.CheckProjectIsOpen(true);
            
            % Get parameter index
            if ischar(ParamNameOrIndex)  % if name is provided
                ParamIndex = this.GetParameterIndexByName(ParamNameOrIndex);
            else % if it is already index
                ParamIndex = ParamNameOrIndex;
                NParams = this.GetNumberOfParameters();
                assert(ParamIndex<=NParams, 'Parameter index (%i) exceeds number of parameters in the project (%i)', ParamIndex, NParams);
            end
            
            ParamExp = this.FProj.invoke('GetParameterSValue', ParamIndex-1);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function CopyAllParametersToMatlabWorkspace(this, Workspace)
            this.CheckProjectIsOpen(true);
            
            if nargin<2, Workspace = 'caller'; end
            
            ValidStrings = {'base', 'caller'};
            Workspace = validatestring(Workspace, ValidStrings, '', 'Workspace');
            
            [ParamNames, ParamValues] = this.GetCurrentValueOfAllParameters();
            for ipar=1:length(ParamNames)
                assignin(Workspace, ParamNames{ipar}, ParamValues(ipar));
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function Solve(this, varargin)
            this.CheckProjectIsOpen(true);
            
            % parse additional arguments
            ValidArgs = {'NumberOfTries', 'SkipIfSolutionExists'};
            iarg = 1;
            funcName = 'Solve';
            NumberOfTries = 1;
            SkipIfSolutionExists = false;
            while iarg<=length(varargin),
                ArgName = validatestring(varargin{iarg}, ValidArgs, funcName, '', iarg);
                switch ArgName,
                    case 'NumberOfTries',
%                         CheckIfValueSpecified(varargin,iarg);
                        NumberOfTries = varargin{iarg+1};
                        validateattributes(NumberOfTries,{'numeric'},{'nonempty','positive','scalar'},funcName,ArgName,iarg+1);
                        iarg = iarg+2;
                    case 'SkipIfSolutionExists',
%                         CheckIfValueSpecified(varargin,iarg);
                        SkipIfSolutionExists = varargin{iarg+1};
                        validateattributes(SkipIfSolutionExists,{'logical'},{'nonempty'},funcName,ArgName,iarg+1);
                        iarg = iarg+2;
                end
            end
            
            if SkipIfSolutionExists,
                [SolutionExists, ParamString] = this.SolutionExistsForCurrentParameterCombination();
                if SolutionExists,
                    this.fprintf('Simulation skipped because a solution exists for the current parameter combination:\n');
                    this.fprintf('  %s\n', ParamString);
                    return
                end
            end
            
            % if we want Y- and/or Z-matrices, activate their automatic calculation; or deactivate; or do nothing  
            this.setCalculateYZMatrices(this.FCalculateYZMatrices);
            
            % start solver depending on its type
            SolverType = this.FProj.invoke('GetSolverType');
            switch SolverType,
                case 'HF Time Domain',
                    this.FSolver = this.FProj.invoke('Solver');
                    this.RunSolver(NumberOfTries, 'Start');
                case 'HF Frequency Domain',
                    this.FSolver = this.FProj.invoke('FDSolver');
                    this.RunSolver(NumberOfTries, 'Start');
                case 'HF Eigenmode',
                    this.FSolver = this.FProj.invoke('EigenmodeSolver');
                    this.RunSolver(NumberOfTries, 'Start');
                case 'HF IntegralEq',
%                     this.FSolver = this.FProj.invoke('FDSolver');
%                     this.RunSolver(NumberOfTries, 'Start');
                    error('Starting the Intergal equation solver is not implemented. It is unclear how to start it properly. Sometimes the Frequency domain solver starts instead...');
                otherwise,
                    error('Unknown solver "%s".',SolverType);
            end
            
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function ResIDs = GetResultIDsFromTreeItem(this, TreeItem)
            this.CheckProjectIsOpen(true);
            
            ResultTree = this.FProj.invoke('Resulttree');
            assert(ResultTree.invoke('DoesTreeItemExist',TreeItem), 'Result tree item "%s" does not exist.', TreeItem);
            ResIDs = ResultTree.invoke('GetResultIDsFromTreeItem',TreeItem);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function ChildTreeItems = GetResultTreeItemChildren(this, ParentTreeItem)
            this.CheckProjectIsOpen(true);
            ChildTreeItems = {};
            
            ResultTree = this.FProj.invoke('Resulttree');
            assert(ResultTree.invoke('DoesTreeItemExist',ParentTreeItem), 'The parent tree item "%s" does not exist.', ParentTreeItem);
            Child = ResultTree.invoke('GetFirstChildName', ParentTreeItem);
            ich = 1;
            while ~isempty(Child),
                ChildTreeItems{ich,1} = Child; %#ok<AGROW>
                Child = ResultTree.invoke('GetNextItemName', Child);
                ich = ich+1;
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [S, Freq, Zref, Info] = GetSParams(this, varargin)
            % [S, Freq, Zref, Info] = Obj.GetSParams(RunIDFilter, iRows, iCols, PortMode, FreqsToGet)
            TreePath = '1D Results\S-Parameters';
            if nargout>=3,
                [S, Freq, Info] = this.GetSYZParams(TreePath, varargin{:});
                Zref = Info.TreeItemImpedance;
            else
                [S, Freq] = this.GetSYZParams(TreePath, varargin{:});
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [Z, Freq, Info] = GetZParams(this, varargin)
            % [S, Freq, Info] = Obj.GetZParams(RunIDFilter, iRows, iCols, PortMode, FreqsToGet)
            TreePath = '1D Results\Z Matrix';
            if nargout>=3,
                [Z, Freq, Info] = this.GetSYZParams(TreePath, varargin{:});
            else
                [Z, Freq] = this.GetSYZParams(TreePath, varargin{:});
            end
        end
        
        % -----------------------------------------------------------------
        % NOTE: Unfortunatelly, we cannot use the CST function "GetParameterCombination" 
        % directly since it has several output arguments, and such
        % functions cannot be called using "invoke" method, and CST has not
        % implemented interface for their COM objects (at least in the version
        % 2019 and below). Therefore a macro-wrapper is used which saves
        % the result of "GetParameterCombination" to a file, and then it is
        % read and processed by Matlab.
        % -----------------------------------------------------------------
        function [ParamNames, ParamValues, ResIDs] = GetParameterCombination(this, RunIDs, varargin)
        % [ParamNames, ParamValues, ResIDs] = obj.GetParameterCombination() - read parameter combination for all RunIDs  
        % [ParamNames, ParamValues, ResIDs] = obj.GetParameterCombination(RunIDs)
        % NOTE: "RunIDs" are NOT RunID filter, so they cannot be negative integers or strings with "~"! 
        
        % We could use RunIDFilter instead of RunIDs, but then we would 
        % need to know all available RunIDs. 
        % Getting them is non-trivial because it requires knowledge of a
        % TreeItem. In case when RunIDs=[] we use S11 TreeItem to get them,
        % but it is better to avoid hard-coded TreeItems, therefore we
        % leave an option of direct specification of RunIDs for which we want
        % to get the parameter combination.
            
            this.CheckProjectIsOpen(true);
            if nargin<2, RunIDs = []; end
            ParamNames = {};  ParamValues = []; 
            
            % define file names
%             MacroDir  = fullfile(tempdir,  'CSTMatlabInterfaceTempFiles');
%             MacroFile = fullfile(MacroDir, 'GetParameterCombinationWrapper.mcr');
%             DataFile  = fullfile(MacroDir, 'GetParameterCombinationData.txt');
%             if ~exist(MacroDir, 'dir'),  mkdir(MacroDir);  end
            [~, MacroFile, DataFile] = this.GetTempDirAndFiles('GetParameterCombination');
            
            % If RunIDs is not specified or empty, we want to get parameters for all result IDs  
            if isempty(RunIDs),
                % Here we use a hardcoded S11 TreeItem to get list of available RunIDs. Hopefully it will not cause any problems in future...  
                TreeItem = '1D Results\S-Parameters\S1,1';  % Tree item from which we will get all Result IDs
                try
                    ResIDs = GetResultIDsFromTreeItem(this, TreeItem);
                catch E,
                    error('%s\nTry to specify RunIDs as input argument if you are sure that results exist.',E.message)
                end
                % Exclude current run, since its results are already present in the Result Navigator with another RunID  
                ResIDs(strcmpi(ResIDs,'3D:RunID:0')) = [];
            else
                [ResIDs, Flags] = this.GetRunIDsStrings(RunIDs);
                assert(all(~Flags), '"RunIDs" for this method are not RunIDFilters, so they cannot be negative or contain "~".');
            end
            
            % get parameters for every result
            tid = tic;  PrintProgress = false;
            NResults = length(ResIDs);
            for ires=1:NResults,
                % select the result
                ResID = ResIDs{ires};
                
                % if it takes more than 3 sec to read, print progress.
                if ~PrintProgress && (toc(tid)>3),  PrintProgress = true;  end
                if PrintProgress,
                    nsym = this.fprintf('Getting parameter combination for "%s" (%i/%i)\n', ResID, ires, NResults);
                end
                
                % write macro file
                this.WriteNamedScriptToFile('GetParameterCombination', MacroFile, {'ResID',ResID; 'DataFile',DataFile});
                % execute it
                this.FProj.invoke('RunScript',MacroFile);
                % read data saved by the macro
                d = importdata(DataFile);
                if iscell(d) && isscalar(d) && strcmpi(d{1},'Parameter combination does not exist.'),
                    error('Parameter combination does not exist for the Result ID "%s".', ResID);
                end
                if ires==1,
                    ParamNames  = d.textdata;
                    NParams = length(ParamNames);
                    ParamValues = nan(NParams, NResults);
                    % if we have not requested parameter values, but just names, return
                    if nargout<2, return; end
                end
                ParamValues(:,ires) = d.data; %#ok<AGROW>
                
                if PrintProgress,
                    this.fprintf(repmat('\b',1,nsym));
                end
            end   
             
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [ResIDs, SolParamNames, SolParamValues, SolResIDs] = FindResultIDsForParameterCombination(this, ParamNames, ParamValues)
            this.CheckProjectIsOpen(true);
            ResIDs = []; 
            
            % if ParamNames is string, make is cell
            if ischar(ParamNames),  ParamNames = {ParamNames};  end
            
            % check input
            assert(iscell(ParamNames), '"ParamNames" must be a cell array with parameter names.');
            assert(isnumeric(ParamValues), '"ParamValues" must be a numerical vector with parameter values.');
            NParams = length(ParamNames);
            assert(NParams==length(ParamValues), 'Number of parameter names in "ParamNames" do not mutch the number of parameter values in "ParamValues".');
            
            % Check that all specified parameters exist
            for ipar=1:NParams,
                this.ParameterExists(ParamNames{ipar}, true);
            end
            
            % Get paramter names and values for all available solutions 
            try % an error will occur if there is no solutions in the project  
                [SolParamNames, SolParamValues, SolResIDs] = this.GetParameterCombination();
            catch
                return
            end

            % Get indices of parameters to check
            Inds = nan(NParams,1);
            for ipar=1:NParams,
                Inds(ipar) = find(strcmpi(SolParamNames,ParamNames{ipar}),1);
            end
            
            if ~iscolumn(ParamValues), ParamValues=ParamValues.'; end
            Dif = bsxfun(@(M1,M2) abs(M1-M2)<=abs(1e-6*ParamValues), SolParamValues(Inds,:), ParamValues);
            ind = all(Dif,1); % index of the solution in the SolParamValues matrix 
            
            % since some RunIDs can be missing (e.g. there was a rebuild error), we need
            % to find the RunID corresponding to the solution index in  SolParamValues matrix 
            numSolResIDs = this.RunIDsStrToNum(SolResIDs);
            ResIDs = numSolResIDs(ind);
            
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [SolutionExists, ParamString, ResID] = SolutionExistsForCurrentParameterCombination(this)
            this.CheckProjectIsOpen(true);
            SolutionExists = false;   ParamString = '';
            
            % Get paramter names and values currently active  
            [ParamNames, ParamValues] = this.GetCurrentValueOfAllParameters();
            ResID = this.FindResultIDsForParameterCombination(ParamNames, ParamValues);
            
            if isempty(ResID),  return;  end            
            assert(length(ResID)==1, 'There are several solutions for the current parameter set?? Bug?');
            SolutionExists = true;

% % %             % Get paramter names and values for all available solutions 
% % %             try % an error will occur if there is no solutions in the project  
% % %                 [SolParamNames, SolParamValues] = this.GetParameterCombination();
% % %                 NSolutions = size(SolParamValues, 2);
% % %                 NParams = length(SolParamNames);
% % %             catch
% % %                 return
% % %             end
% % %             
% % %             % Get paramter names and values currently active  
% % %             [ParamNames, ParamValues] = this.GetCurrentValueOfAllParameters();   
% % %             % do some checking, which should not happen, but just in case 
% % %             assert(isequal(ParamNames,SolParamNames), '"ParamNames" and "SolParamNames" are not equal. Unexpected.');
% % %             
% % %             % Check if there is a solution exists for the current parameters set 
% % %             iSolution = [];
% % %             for isol=1:NSolutions,
% % %                 if isequal(ParamValues,SolParamValues(:,isol)),
% % %                     SolutionExists = true;
% % %                     iSolution = isol;
% % %                     break
% % %                 end
% % %             end
            
            % Form a string containing parameter name-value pairs 
            if SolutionExists && nargout>=2,
                ParamString = sprintf('Solution RunID %i:  ',ResID);
                NParams = length(ParamNames);
                for ipar=1:NParams,
                    if ipar~=NParams, sep = ';  '; else, sep = ''; end
                    ParamString = sprintf('%s%s = %g%s',ParamString,ParamNames{ipar},ParamValues(ipar),sep);
                end
            end
            
        end
        
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function ExportTouchstone(this, FileName, NormalizationImpedance, NumberOfSamples, FreqRange)
        % Obj.ExportTouchstone(FileName, NormalizationImpedance, NumberOfSamples, FreqRange)    
        % "NormalizationImpedance" may be [], then S-params will not be normalized.
        % "NumberOfSamples": default value is 1001.
        % "FreqRange" may be string 'Full' (= export all available frequency range) or 1x2 numeric vector ([f_min f_max]).  
        % NOTE: CST adds to the FileName extension corresponding to the number of ports (e.g. ".s2p", "s25p" etc).  
            
            this.CheckProjectIsOpen(true);
            narginchk(2,5);
            
            if nargin<3, NormalizationImpedance = 50; end
            if nargin<4, NumberOfSamples = 1001; end
            if nargin<5, FreqRange = 'Full'; end
            
            funcName = 'ExportTouchstone';
            if ~isempty(NormalizationImpedance),
                validateattributes(NormalizationImpedance,{'numeric'},{'positive','scalar'},funcName,'NormalizationImpedance',2);
            end
            validateattributes(NumberOfSamples,{'numeric'},{'nonempty','positive','scalar'},funcName,'NumberOfSamples',3);
            validateattributes(FreqRange,{'numeric','char'},{'nonempty'},funcName,'FreqRange',3);
            assert(isnumeric(FreqRange)||strcmpi(FreqRange,'Full'), '"FreqRange" must be 1x2 numeric vector ([f_min f_max]) of string ''Full''.');
            
            FileName = TCSTInterface.GetFullPath(FileName);
            
            % Since CST adds the file extension, remove it from the FileName if it is present  
            FileName = regexprep(FileName, '.[sS]\d+[pP]\s*$', '');
            
            objTS = this.FProj.invoke('TOUCHSTONE');
%             objTS = this.FProj.invoke('CallByName','TOUCHSTONE');
            objTS.invoke('Reset');
            objTS.invoke('FileName', FileName);
            if ~isempty(NormalizationImpedance)
                objTS.invoke('Impedance',NormalizationImpedance);
                objTS.invoke('Renormalize',true);
            else
                objTS.invoke('Renormalize',false);
            end
            objTS.invoke('UseARResults',false);
            objTS.invoke('SetNSamples',NumberOfSamples);
            if ischar(FreqRange),
                objTS.invoke('FrequencyRange', 'Full');
            else
                objTS.invoke('Fmin', FreqRange(1));
                objTS.invoke('Fmax', FreqRange(2));
                objTS.invoke('FrequencyRange', 'Limited');
            end
            objTS.invoke('Write');
            
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [Field, Info] = ExportFarField(this, OutputFileName, SaveBroadband, ThetaPhiStepDeg, NormalizationType, SaveToMat)
            this.CheckProjectIsOpen(true);
            narginchk(1,6);
            
            if nargin<2, OutputFileName = []; end 
            assert(~isempty(OutputFileName) || nargout>0, 'No "OutputFileName" is supplied. An output argument is expected in this case.');
            if nargin<3 || isempty(SaveBroadband),     SaveBroadband = false;      end  
            if nargin<4 || isempty(ThetaPhiStepDeg),   ThetaPhiStepDeg = [5 5];    end            
            if nargin<5 || isempty(NormalizationType), NormalizationType = 'None'; end
            if nargin<6 || isempty(SaveToMat),         SaveToMat = false;          end
            
            % if we want to save to files
            if ~isempty(OutputFileName),
                % Prepare FileName
                OutputFileName = TCSTInterface.GetFullPath(OutputFileName);
                [FileDir, FileName, FileExt] = fileparts(OutputFileName);
                FileExt = lower(FileExt);
                assert(~isempty(intersect(FileExt,{'.ffs','.mat'})), 'Only file extensions ".ffs" and ".mat" are supported.');
                switch FileExt,
                    % this means that we want FFS files (and .mat if SaveToFieldStruct=true)  
                    case '.ffs',
                        DeleteFFSAfterReading = false;
                    % this means that we do not need FFS files; delete them after reading to the Field structure 
                    case '.mat',
                        DeleteFFSAfterReading = true;
                        SaveToMat = true; % override
                    otherwise, error('Bug: Unexpected "FileExt".');
                end
                
            % if we do NOT want to save to files
            else
% % %                 FileDir = fullfile(tempdir, 'CSTMatlabInterfaceTempFiles');
                FileDir = this.GetTempDirAndFiles();
                FileName = 'TempFarField';
                FileExt = '.ffs';
                DeleteFFSAfterReading = true;
            end
            
            if ~exist(FileDir,'dir'), mkdir(FileDir); end
            
            % Get names of TreeItems with the far fields and remove 'Farfield Cuts' from there 
            sTreeItems = this.GetResultTreeItemChildren('Farfields');
            sTreeItems = setdiff(sTreeItems, 'Farfields\Farfield Cuts', 'stable');
            assert(~isempty(sTreeItems), 'CSTInterface:Export:NoFarFields', 'There are no far fields calculated.');
            NFarFields = length(sTreeItems);
            
%             % Get ResultTree object from the project
%             ResultTree = this.FProj.invoke('Resulttree');
            
            % ------------ Export far fields to FFS files --------------
            % file names of the saved FFS files and corresponding excitation indices 
            SavedFFSFiles = cell(NFarFields,1);   iSvdFfs = 1; 
            SavedExcitationInd = nan(NFarFields,1);
            % export each far field
            for iti=1:NFarFields
                sTreeItem = sTreeItems{iti};
                
                % get frequency and excitation index
                NumberMatch = '(?:[-+]?\d*\.?\d+)(?:[eE]([-+]?\d+))?';
%                 RegExpPattern = ['.*\(f\s?=\s?', '([(?:' NumberMatch ')|(?:broadband)])', '\)\s*\[', '(\d+)', '\]'];
                RegExpPattern = ['.*\((?:f\s?=\s?)?', '((?:' NumberMatch ')|(?:broadband))', '\)\s*\[', '(\d+)', '\]'];
                d = regexp(sTreeItem, RegExpPattern, 'tokens');
                assert(~isempty(d) && length(d{1})==2, 'Cannot get frequency and excitation index from the tree item name "%s".', sTreeItem);
                strFreq = d{1}{1};
                iExc = str2double(d{1}{2});
                IsBroadband = strcmpi(strFreq,'broadband');
                
                % skip this TreeItem if it doesn't correspond requested "broadband" criteria 
                if SaveBroadband && ~IsBroadband,   continue;    end
                if ~SaveBroadband && IsBroadband,   continue;    end
                
                % form file name for the current FFS file 
                FileSuffix = strrep( sprintf('_%s_Exc%i', strFreq, iExc), '.', 'p');
                CurrentFileName = fullfile(FileDir, [FileName FileSuffix '.ffs']);
                
                % select one far field result
                this.FProj.invoke('SelectTreeItem',sTreeItem);
                % plot it
                FarfieldPlot = this.FProj.invoke('FarfieldPlot');
                FarfieldPlot.invoke('Reset');
                FarfieldPlot.invoke('Plottype', '3d');
                FarfieldPlot.invoke('SetLockSteps',false);
                FarfieldPlot.invoke('Step',  ThetaPhiStepDeg(1));
                FarfieldPlot.invoke('Step2', ThetaPhiStepDeg(2));
                FarfieldPlot.invoke('Plot');
                % export
                if SaveBroadband,
                    FarfieldPlot.invoke('ASCIIExportAsBroadbandSource',CurrentFileName);
                else
                    FarfieldPlot.invoke('ASCIIExportAsSource',CurrentFileName);
                end
                
                % add the file name to the "saved" list
                SavedFFSFiles{iSvdFfs} = CurrentFileName;   
                SavedExcitationInd(iSvdFfs) = iExc;
                iSvdFfs = iSvdFfs+1;
            end
            % remove empty cells
            ind = isnan(SavedExcitationInd);
            SavedFFSFiles(ind)=[];
            SavedExcitationInd(ind)=[];
            NFFSFiles = length(SavedFFSFiles);
            
            if SaveBroadband,
                assert(NFFSFiles>0, 'There are no broadband far fields found in the CST project results.');
            else
                assert(NFFSFiles>0, 'There are no single-frequency far fields found in the CST project results. Are there broadband far-field monitors? Set SaveBroadband=true is this case.');
            end
            
            % -------------- Read FFS to Field structure ---------------   
            if nargout>0 || SaveToMat,
                pause(0.5);
                for ifl=1:NFFSFiles,
                    % read ffs
                    [F, Info(ifl)] = TCSTInterface.ReadFarFieldSourceFile(SavedFFSFiles{ifl}, [], NormalizationType);
                    % if this is 1st file, save whole field struct 
                    if ifl==1,
                        Field = F;
                    end
                    % excitation to which this FFS file corresponds to 
                    iExc = SavedExcitationInd(ifl);
                    if SaveBroadband,
                        Field.E(:,:,:,iExc,:) = F.E;
                    else
                        % find freq index to which F.E corresponds and save the field   
                        iFr = find(abs(Field.Freq-F.Freq)<1000*eps(F.Freq),1);
                        if isempty(iFr), % new frequency
                            iFr = length(Field.Freq)+1;
                        end
                        Field.E(:,:,:,iExc,iFr) = F.E;
                        Field.Freq(iFr) = F.Freq;
                    end
                    
                end
            end
            
            % -------------- Save MAT file if requested ---------------- 
            if SaveToMat,
                MatFileName = fullfile(FileDir, [FileName '.mat']);
                save(MatFileName, 'Field');
            end
            
            % ----------- Delete FFS if they are not needed ------------ 
            if DeleteFFSAfterReading,
                delete(SavedFFSFiles{:});
            end

        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [Res, ResultTree] = TreeItemExists(this, TreeItem, ThrowErrorIfNotExist)
            if nargin<3, ThrowErrorIfNotExist = false; end
            this.CheckProjectIsOpen(true);
            ResultTree = this.FProj.invoke('Resulttree');
            Res = ResultTree.invoke('DoesTreeItemExist',TreeItem);
            assert(Res || ~ThrowErrorIfNotExist, 'There is no tree item "%s" in the project.', TreeItem);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [Xout,Yout, Imp, RunIDsToProcess, Info] = Get1DResultFromTreeItem(this, TreeItem, RunIDFilter, QueryX)    
            this.CheckProjectIsOpen(true);
            if nargin<3, RunIDFilter = []; end
            if nargin<4, QueryX = []; end % read all X
            Imp = [];
            
            % Get ResultTree object from the project
            this.TreeItemExists(TreeItem, true);
            ResultTree = this.FProj.invoke('Resulttree');
            
            % Filter ResultIDs if specified (runs for e.g. parametric sweep). 
            RunIDs = this.GetResultIDsFromTreeItem(TreeItem);
            RunIDsToProcess = this.GetFilteredRunIDs(RunIDs, RunIDFilter);
            NResults = length(RunIDsToProcess);
            
            % Get X
            [~, ~, Info] = this.Get1DResultXFor1stRunID(TreeItem, QueryX);  
            objRes = ResultTree.invoke('GetResultFromTreeItem', TreeItem, RunIDsToProcess{1});
            X = objRes.invoke('GetArray','x');    if isrow(X), X = X.'; end
            Xout = this.FindQueriedX(X, QueryX, true);
            NX = length(Xout);
            this.fprintf('Data for %i points will be read: X = %s\n', NX, TCSTInterface.VectorToString(Xout));

%             % Filter ResultIDs if specified (runs for e.g. parametric sweep). 
%             RunIDsToProcess = this.GetFilteredRunIDs(RunIDs, RunIDFilter);
%             NResults = length(RunIDsToProcess);
            
            Yout = nan(NX, NResults);
            if Info.TreeItemHasImpedance,
                Imp = nan(NX, NResults);
            end
            for iRes=1:NResults,
                RunID = RunIDsToProcess{iRes};
                this.fprintf('Reading "%s", result id "%s"...\n', TreeItem, RunID);
                
                % Get Result1D object and check its type
                objRes = ResultTree.invoke('GetResultFromTreeItem', TreeItem, RunID);
                ResType = objRes.invoke('GetResultObjectType');
                assert(~isempty(intersect(ResType,{'1D','1DC'})), 'Tree item "%s" contains unsupported result type "%s".',TreeItem,ResType);
                IsComplexData = strcmpi(ResType,'1DC');
                       
                % Get X
                X = objRes.invoke('GetArray','x');
                if isrow(X), X = X.'; end
                % Check that result for this RunID has same X vector
                this.CheckXIsSame(Xout, X, RunID, QueryX);
                
                % Get Y
                if IsComplexData,
                    Y = objRes.invoke('GetArray','yre') + 1i*objRes.invoke('GetArray','yim');
                else
                    Y = objRes.invoke('GetArray','y');
                end

                % get Y values (interpolation will be performed if requested and needed) 
                [~, iX] = this.FindQueriedX(X, QueryX, false);
                Yout(:,iRes) = this.GetIndexedY(X, Y, iX, QueryX);
                
                % "nargout>=3" : not always Info is requested. Save some reading time and do not read TreeItemHasImpedance in this case.   
                if nargout>=3 && Info.TreeItemHasImpedance,
                    objTreeItemImpedance = ResultTree.invoke('GetImpedanceResultFromTreeItem', TreeItem, RunID);
                    ResType1 = objTreeItemImpedance.invoke('GetResultObjectType');
                    IsComplexData1 = strcmpi(ResType1,'1DC');
                    
                    if IsComplexData1,
                        Z = objTreeItemImpedance.invoke('GetArray','yre') + 1i*objTreeItemImpedance.invoke('GetArray','yim');
                    else
                        Z = objTreeItemImpedance.invoke('GetArray','y');
                    end
                    Imp(:,iRes) = this.GetIndexedY(X, Z, iX, QueryX); %#ok<AGROW> It is not growing (already initialized)
                end

            end
            
            % If all impedances are same, keep only one of them
            if nargout>=3 && Info.TreeItemHasImpedance && ( all(abs(Imp(:)-Imp(1))<100*eps(Imp(1))) ),
                Imp = Imp(1);
            end
            
            
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function D = ReadParametricResults(this, varargin)
            this.CheckProjectIsOpen(true);
            
            % Process arguments
            narginchk(2,inf); % minimum 1 argument must be supplied
            ResultsToRead = varargin{1};   if ischar(ResultsToRead)||isstring(ResultsToRead), ResultsToRead = {ResultsToRead}; end
            if nargin<3, FilterRunIDs = nan;  % exclude current run
            else, FilterRunIDs = varargin{2};  end
            ValidArgs = { ...
                'VarNames', ...
                'CacheFile', ...
                'SplitParameters', ...
                'QueryX' ...
            };
            funcName = 'ReadParametricResults';
            iarg = 3;
            VarNames = {};
            CacheFile = [];
            QueryX = [];
            SplitParameters = false;
            while iarg<=length(varargin),
                ArgName = validatestring(varargin{iarg}, ValidArgs, funcName, '', iarg);
                switch ArgName,
                    case 'VarNames'
                        this.CheckIfValueSpecified(varargin,iarg, {'char','cell'});
                        VarNames = varargin{iarg+1};
                        if ischar(VarNames)||isstring(VarNames), VarNames = {VarNames}; end
                        iarg = iarg+2;
                    case 'CacheFile'
                        this.CheckIfValueSpecified(varargin,iarg, {'char'});
                        CacheFile = TCSTInterface.GetFullPath( varargin{iarg+1} );
                        iarg = iarg+2;
                    case 'SplitParameters',
                        this.CheckIfValueSpecified(varargin,iarg, {'logical'});
                        SplitParameters = varargin{iarg+1};
                        iarg = iarg+2;  
                    case 'QueryX',
                        this.CheckIfValueSpecified(varargin,iarg, {'numeric'});
                        QueryX = varargin{iarg+1};
                        iarg = iarg+2;
                end
            end
            NResultsToRead = length(ResultsToRead);
            
            % Get variable names if not present
            % TODO! Result - VarNames must be defined for each ResultsToRead
            assert(isempty(VarNames) || (length(VarNames)==NResultsToRead), ...
                '"VarNames" must have the same length as "ResultsToRead" (%i), or empty. Actual size of "VarNames" is %i.', ...
                NResultsToRead, length(VarNames));
            if isempty(VarNames),
                VarNames = cell(1,NResultsToRead);
            end
            
            % Check that all specified ResultsToRead present, and find available RunIDs  
            AvailableRunIDs = {};
            [ResultsToReadExceptSZ, ia] = setdiff(ResultsToRead, {'S','Z'}); % except S and Z, they are special cases
            VarNamesExceptSZ = VarNames(ia);
            for ir=1:length(ResultsToReadExceptSZ),
                TreeItem = ResultsToReadExceptSZ{ir};
                this.TreeItemExists(TreeItem, true);
                if isempty(AvailableRunIDs),
                    AvailableRunIDs = this.GetResultIDsFromTreeItem(TreeItem);
                end
            end
            if isempty(AvailableRunIDs) && any(strcmpi(ResultsToRead,'Z')),
                TreePath = '1D Results\Z Matrix';
                ResultTree = this.FProj.invoke('Resulttree');
                assert(ResultTree.invoke('DoesTreeItemExist',TreePath), 'There is no result path "%s".', TreePath);
                TreeItem = ResultTree.invoke('GetFirstChildName', TreePath);
                AvailableRunIDs = this.GetResultIDsFromTreeItem(TreeItem);
            end
            if isempty(AvailableRunIDs) && any(strcmpi(ResultsToRead,'S')),
                TreePath = '1D Results\S-Parameters';
                ResultTree = this.FProj.invoke('Resulttree');
                assert(ResultTree.invoke('DoesTreeItemExist',TreePath), 'There is no result path "%s".', TreePath);
                TreeItem = ResultTree.invoke('GetFirstChildName', TreePath);
                AvailableRunIDs = this.GetResultIDsFromTreeItem(TreeItem);
            end
            assert(~SplitParameters||~isempty(AvailableRunIDs), 'Cannot split parameters: "AvailableRunIDs" were not found. Bug?');

            % Create storage for the data and its argument (e.g. frequency and/or parameters)  
            D = TResultsStorage();
            D.FDoNotDisplayColumns = {'Coefficient','Label','Legend','Title','UserData'};
            
            % If CacheFile specified, load data from it if present
            if ~isempty(CacheFile) && isfile(CacheFile),
                load(CacheFile, 'D'); % D will be overwritten
                % Remove the loaded results from the query ResultsToRead
                ExistingVarNames = D.GetDependentVariables();
                [VarNames, ia] = setdiff(VarNames, ExistingVarNames);
                ResultsToRead = ResultsToRead(ia);
                % if all requested results are present - return, since we already loaded our data D  
                if isempty(VarNames), 
                    this.fprintf('ReadParametricResults: All requested results have been loaded from the cache file "%s".\n', CacheFile);
                    return;  
                end
                s = sprintf('"%s", ',ResultsToRead{:});
                this.fprintf('ReadParametricResults: Some of the requested results have been loaded from the cache file "%s".\n', CacheFile);
                this.fprintf('ReadParametricResults: The following results will be now read:\n');
                this.fprintf('  %s\n',s(1:end-2));
            end
            % At this point we have to read remaining (after loading from the cache) ResultsToRead 
            % and store them to D with corresponding names in VarNames
            
            % If we want to split parameters to different dimensions, get parameter     
            % combination for specified RunIDs
            if SplitParameters,
                
                % get parameter combination
                RunIDs = this.GetFilteredRunIDs(AvailableRunIDs, FilterRunIDs);
                [ParamNames, ParamValues] = this.GetParameterCombination(RunIDs);
                if ~isrow(ParamNames),  ParamNames=ParamNames.'; end
                
                % exclude non-changing parameters
                ind = ~all(ParamValues==ParamValues(:,1),2);
                ParamNames = ParamNames(ind);
                ParamValues = ParamValues(ind,:);
                
                % exclude dependent parameters
                NParams = length(ParamNames);
                iIndependentPars = false(NParams,1);
                for ip=1:NParams,
                    Expr = this.GetParameterExpression(ParamNames{ip});
                    % if parameter expression contains only symbols in set [0-9 . ,], we consider it independent  
                    iIndependentPars(ip) = all( (Expr<=int8('9') & Expr>=int8('0')) | Expr==int8('.') | Expr==int8(',') );
                end
                ParamNames = ParamNames(iIndependentPars);
                ParamValues = ParamValues(iIndependentPars,:);
                
                % Find unique parameter values
                NParams = length(ParamNames);
                ParamValuesUnique = cell(NParams,1);
                for ip=1:NParams,
                    ParamValuesUnique{ip} = uniquetol(ParamValues(ip,:), 1e-9);
                end
%                 NParamValuesUnique = cellfun(@(c)length(c), ParamValuesUnique);
                
                % save to the result storage
                for ip=1:NParams,
                    if ~isprop(D,ParamNames{ip}),
                        D.AddVariable(ParamNames{ip},ParamValuesUnique{ip});
                    end
                end
                
                % Form matrix ParamIndices containing indices of parameters in ParamValuesUnique  
                % for each parameter (1st dim) and RunID (2nd dim).
                % E.g.  ParamNames={'par1';'par2'}  and  ParamValuesUnique={[4 5],[7 8 9]}  
                % means that there are 2 changing parameters in the CST project: par1 and par2,  
                % which have values par1=[4 5] and par2=[7 8 9]. Suppose we did a parametric  
                % sweep in the project with these parameters and therefore we have 6 RunIDs. 
                % In this case ParamValues=[4 4 4 5 5 5;  7 8 9 7 8 9], in
                % which 1st dim is parameter index and 2nd dim is RunID.
                % "ParamIndices" has same size as "ParamValues" and set
                % correspondence between "ParamValues" and "ParamValuesUnique".  
                % In this exaple ParamIndices=[1 1 1 2 2 2;  1 2 3 1 2 3].
                % "ParamIndices" is used later for resaping result's RunID dimension. 
                [NParams, NRunIDs] = size(ParamValues);
                ParamIndices = nan(NParams,NRunIDs);
                for ir = 1:NRunIDs,
                    for ip=1:NParams,
                        n = find(ismembertol(ParamValuesUnique{ip}, ParamValues(ip,ir), 1e-6), 1);
                        assert(~isempty(n), 'Value %g of parameter %i ("%s") was not found. Bug?', ParamValues(ip,ir), ip, ParamNames{ip});
                        assert(isnan(ParamIndices(ip,ir)), 'ParamIndices(ip,ic) is not NaN, meaning that . Bug?');
                        ParamIndices(ip,ir) = n;
                    end
                end

            end
            
            % ------- SPECIAL CASE: Z -------
            iRes = find(strcmpi(ResultsToRead,'Z'),1);
            if ~isempty(iRes),
                % get Z-matrix
                [Zraw, ZFreq, ZInfo] = this.GetZParams(FilterRunIDs,[],[],[],QueryX);
                Nm = size(Zraw,1);
                Nn = size(Zraw,2);
                
                % Get function name
                if isempty(VarNames{iRes}),  FunName = 'Z';
                else,  FunName = VarNames{iRes};
                end
                
                % Save Z-matrix to the data storage, reshape RunID dimension by parameters if requested 
                D = localStoreVariable(D, FunName,Zraw, {'m','n','ZFreq'},{1:Nm,1:Nn,ZFreq}, ZInfo.ResultIDs, 'Description','Full Z-matrix');
            end
            % -------------------------------
            
            % ------- SPECIAL CASE: S -------
            iRes = find(strcmpi(ResultsToRead,'S'),1);
            if ~isempty(iRes),
                % get S-matrix
                [Sraw, SFreq, Zref, SInfo] = this.GetSParams(FilterRunIDs,[],[],[],QueryX);
                Nm = size(Sraw,1);
                Nn = size(Sraw,2);
                
                % Get function name
                if isempty(VarNames{iRes}),  FunName = 'S';
                else,  FunName = VarNames{iRes};
                end
                
                % Save Zref
                D.AddVariable('Zref',Zref, 'Description','Reference impedance for the S-parameters');
                % Save Z-matrix to the data storage, reshape RunID dimension by parameters if requested 
                D = localStoreVariable(D, FunName,Sraw, {'m','n','SFreq'},{1:Nm,1:Nn,SFreq}, SInfo.ResultIDs, 'Description','Full S-matrix');
            end
            % -------------------------------
            
            % ------ ALL OTHER RESULTS ------
            for ir=1:length(ResultsToReadExceptSZ),
                % Get TreeItem's data
                TreeItem = ResultsToReadExceptSZ{ir};
                [ArgVal,FunVal, ~, ReadRunIDs, Info] = this.Get1DResultFromTreeItem(TreeItem, FilterRunIDs, QueryX);
                
                % Get argument name from CST plot label
                ArgName = matlab.lang.makeValidName(Info.XLabel);
                % Get function name
                if isempty(VarNamesExceptSZ{ir}),
                    FunName = matlab.lang.makeValidName(Info.Title);
                else
                    FunName = VarNamesExceptSZ{ir};
                end
                
                % Save Z-matrix to the data storage, reshape RunID dimension by parameters if requested 
                D = localStoreVariable(D, FunName,FunVal, {ArgName},{ArgVal}, ReadRunIDs, 'Description',['Obtained from "' TreeItem '"']);
            end
            % -------------------------------
            
            % If cache file is specified, save data to it
            if ~isempty(CacheFile),
                CacheFileDir = fileparts( CacheFile );
                if ~TCSTInterface.isfolder(CacheFileDir), mkdir(CacheFileDir); end
                save(CacheFile, 'D');
                this.fprintf('The data has been saved to file "%s".\n', CacheFile);
            end
            
            
            % ----------- FUNCTION ADDING THE FunVal TO THE STORAGE D --------------------
            function D = localStoreVariable(D, FunName,FunVal, ArgNames,ArgVals, ReadRunIDs, varargin)
                
                ArgDimsLen = cellfun(@(c)length(c), ArgVals, 'UniformOutput',false);
                if ~isrow(ArgDimsLen),  ArgDimsLen=ArgDimsLen.'; end
                ArgDimInds = cellfun(@(c)1:c, ArgDimsLen, 'UniformOutput',false);
                
                % Store the function's arguments
                for iar=1:length(ArgNames),
                    ArgNm = ArgNames{iar};
                    ArgVl = ArgVals{iar};
                    % Check if this argument is already stored.
                    % It can happen that the argument with same name is
                    % already stored in D (while processing previous
                    % result). We do not need to store it again IF it has
                    % same values. However, in some cases it can have
                    % different values. E.g. the radiation efficiency was
                    % computed for 3 frequecies (depends on field monitors
                    % defined), and S11 was computed for 1001 freqs, but
                    % both results have same name Frequency_GHz. We need to
                    % check for a such potential problem, and save the
                    % argument with different name.
                    while true,
                        % if the argument value is not stored yet, store it 
                        if ~isprop(D,ArgNm),
                            if strcmpi(ArgNm,'m') || strcmpi(ArgNm,'n'), ArgDecr = 'Port index';
                            else, ArgDecr = ''; 
                            end
                            D.AddVariable(ArgNm,ArgVl, 'Description',ArgDecr);
                            ArgNames{iar} = ArgNm;
                            break
                        % if the argument with this name already present, check if it has same values     
                        else
                            % Check if stored arg values are the same as ArgVl
                            ValsEqual = ...
                                   ( isa(D.(ArgNm),class(ArgVl)) ) && ... values have the same type
                                   ( numel(D.(ArgNm)) == numel(ArgVl) ) && ...  they have same number of elements ...
                                   ( all(all( abs(D.(ArgNm)-ArgVl) < 1e-6*min(abs(D.(ArgNm)),abs(ArgVl)) )) ); % ... and values are same within 1e-6 tolerance
                            % if they are NOT the same, change name and check again
                            if ~ValsEqual,
                                ArgNm = matlab.lang.makeUniqueStrings(ArgNm, ArgNm);
                            % if they are same, then we already have this argument stored in D, break the checking loop  
                            else
                                ArgNames{iar} = ArgNm;
                                break
                            end
                        end
                    end
                end
                
                % If we want to reshape the function by parameters  
                if SplitParameters,
                    % cell array with length of each parameter. Used to initialize the reshaped function. 
                    ParDimsLen = cellfun(@(c)length(c),ParamValuesUnique, 'UniformOutput',false);
                    if ~isrow(ParDimsLen),  ParDimsLen=ParDimsLen.'; end
                    
                    % reshape the function
                    FunDimsLen = [ArgDimsLen ParDimsLen];
                    FunReshaped = nan(FunDimsLen{:});
                    for irid=1:NRunIDs
                        % indexing CST parameters
                        ind = num2cell(ParamIndices(:,irid).');
                        FunReshaped(ArgDimInds{:},ind{:}) = FunVal(ArgDimInds{:},irid);
                    end
                    
                    % store the function values
                    D.AddVariable(FunName,FunReshaped, 'Arguments',[ArgNames,ParamNames], varargin{:});
                    
                % If we want just to save the function with RunID in the last dimension 
                else
                    
                    if ~isprop(D,'RunID'),
                        % Get RunIDs as numbers
%                         tmp = cellfun(@(c)strsplit(c,':'), ReadRunIDs, 'UniformOutput',false);
%                         locRunIDs = cellfun(@(c)str2double(c{3}), tmp);
                        locRunIDs = TCSTInterface.RunIDsStrToNum(ReadRunIDs);
                        % ... and store them
                        D.AddVariable('RunID',locRunIDs, 'Units','-', 'Description','Run ID in the CST Result Navigator');
                    end
                    
                    % store the function values
                    D.AddVariable(FunName,FunVal, 'Arguments',[ArgNames,{'RunID'}], varargin{:});
                    
                end

            end
            % ------------------------------------------------------------------------
            
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [MonitorNames, MonitorInfos] = GetMonitorNames(this)
            this.CheckProjectIsOpen(true);
            MonObj = this.FProj.invoke('Monitor');
            NMons = MonObj.invoke('GetNumberOfMonitors');
            MonitorNames = cell(NMons,1);
            for im=NMons:-1:1  % we run index backwards to preallocate the structure MonitorInfos
                MonitorNames{im} = MonObj.invoke('GetMonitorNameFromIndex',im-1);
                if nargout>=2,
                    s.MonitorType      = MonObj.invoke('GetMonitorTypeFromIndex',im-1);
                    s.MonitorDomain    = MonObj.invoke('GetMonitorDomainFromIndex',im-1);
                    s.MonitorFrequency = [];
                    s.MonitorTStart    = [];
                    s.MonitorTStep     = [];
                    s.MonitorTEnd      = [];
                    switch lower(s.MonitorDomain)
                        case 'frequency',
                            s.MonitorFrequency = MonObj.invoke('GetMonitorFrequencyFromIndex',im-1);
                        case 'time',
                            s.MonitorTStart    = MonObj.invoke('GetMonitorTstartFromIndex',im-1);
                            s.MonitorTStep     = MonObj.invoke('GetMonitorTstepFromIndex',im-1);
                            s.MonitorTEnd      = MonObj.invoke('GetMonitorTendFromIndex',im-1);
                        case 'static',
                            warning('TCSTInterface:StaticMonitor', 'TO DO: What is "static" monitor domain? What infor can we get for it?');
                        otherwise,
                            error('Monitor "%s" has an unsupported domain "%s".', MonitorNames{im}, s.MonitorDomain);
                    end
                    MonitorInfos(im) = s;
                end
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function Res = MonitorExists(this, MonitorName)
            this.CheckProjectIsOpen(true);
            Res = false;
            MonitorNames = this.GetMonitorNames();
            for im=1:length(MonitorNames)
                Res = strcmp(MonitorNames,MonitorName);
                if Res, break; end
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function DeleteAllMonitors(this)
            this.CheckProjectIsOpen(true);
            MonObj = this.FProj.invoke('Monitor');
            MonitorNames = this.GetMonitorNames();
            for im=1:length(MonitorNames)
                MonObj.invoke('Delete',MonitorNames{im});
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function AddMonitorFreqDomain(this, FieldType, Freq)
            this.CheckProjectIsOpen(true);
            ValidFieldTypes = {'Efield','Hfield','Powerflow','Current','Powerloss','Eenergy','Henergy','Farfield','Fieldsource','Spacecharge','Particlecurrentdensity'};
            FieldType = validatestring(FieldType, ValidFieldTypes, 'AddMonitorFreqDomain', 'FieldType');
            
            switch lower(FieldType)
                case 'efield'
                    Name = sprintf('e-field (f=%g)', Freq);
                case 'hfield'
                    Name = sprintf('h-field (f=%g)', Freq);
                otherwise
                    Name = sprintf('%s (f=%g)', lower(FieldType), Freq);
            end
            
            MonObj = this.FProj.invoke('Monitor');
            MonObj.invoke('Reset');
            MonObj.invoke('Name',         Name);
            MonObj.invoke('Dimension',    'volume');
            MonObj.invoke('Domain',       'frequency');
            MonObj.invoke('FieldType',    FieldType);
            MonObj.invoke('Frequency',    Freq);
            MonObj.invoke('UseSubVolume', false);
            MonObj.invoke('Create');
        end
        
        % -----------------------------------------------------------------
        % It is just a wrapper to GetAllObjectNames
        % -----------------------------------------------------------------
        function ComponentNames = GetComponentNames(this)
            this.CheckProjectIsOpen(true);
            [~,ComponentNames] = this.GetAllObjectNames();
            ComponentNames = unique(ComponentNames);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [ObjectFullNames, ObjectComponents, ObjectNames] = GetAllObjectNames(this, varargin)
            this.CheckProjectIsOpen(true);
            
            % Process arguments
            ValidArgs = { ...
                'Component', ... keep only objects within these specified components
                'Include' ...    keep only objects whos full name contains the specified sub-strings 
                'Exclude' ...    excludes the object whos full name contains the specified sub-strings 
            };
            funcName = 'GetAllObjectNames';
            iarg = 1;
            OPT.Component = {};
            OPT.Include   = {};
            OPT.Exclude   = {};
            while iarg<=length(varargin),
                ArgName = validatestring(varargin{iarg}, ValidArgs, funcName, '', iarg);
                this.CheckIfValueSpecified(varargin,iarg, {'char','cell'});
                OPT.(ArgName) = varargin{iarg+1};
                if ~iscell(OPT.(ArgName)), OPT.(ArgName) = {OPT.(ArgName)}; end
                iarg = iarg+2;
            end

            % Get all object full names (which includes their components), 
            % as well as object's components and names separately  
            Solid = this.FProj.invoke('Solid');
            NObjs = Solid.invoke('GetNumberOfShapes');
            ObjectFullNames  = cell(NObjs,1);
            ObjectComponents = cell(NObjs,1);
            ObjectNames      = cell(NObjs,1);
            for iobj = 1:NObjs
                ObjectFullNames{iobj} = Solid.invoke('GetNameOfShapeFromIndex',iobj-1);
                d = strsplit(ObjectFullNames{iobj},':');
                ObjectComponents{iobj} = d{1};
                ObjectNames{iobj}      = d{2};
            end
            
            % Below we apply filters, if any are specified.
            ind = true(NObjs,1);
            
            % If Component name(s) is/are specified, keep only objects within it/them
            if ~isempty(OPT.Component),
                ind = ind & ismember(ObjectComponents, OPT.Component);
            end
            
            % "Include" filter
            if ~isempty(OPT.Include),
                ind = ind & contains(ObjectFullNames, OPT.Include);
            end
            
            % "Exclude" filter
            if ~isempty(OPT.Exclude),
                ind = ind & ~contains(ObjectFullNames, OPT.Exclude);
            end
            
            ObjectFullNames  = ObjectFullNames(ind);
            ObjectComponents = ObjectComponents(ind);
            ObjectNames      = ObjectNames(ind);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function Res = ObjectExist(this, ObjectFullNames, ThrowErrorIfNot)
            this.CheckProjectIsOpen(true);
            assert(ischar(ObjectFullNames)||iscell(ObjectFullNames), '"ObjectFullName" must be a string or cell array of strings.');
            if ~iscell(ObjectFullNames), ObjectFullNames = {ObjectFullNames}; end
            if nargin<3, ThrowErrorIfNot = false; end
            
            Solid = this.FProj.invoke('Solid');
            NObjs = length(ObjectFullNames);
            Res = false(NObjs,1);
            for iobj=1:NObjs
                Res(iobj) = Solid.invoke('DoesExist',ObjectFullNames{iobj});
                assert(~ThrowErrorIfNot||Res(iobj), 'Object "%s" doesn''t exist in the project. Does its name include the component name (the format is "Path/ComponentName:ObjectName")?', ObjectFullNames{iobj});
            end
            
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [Vol, ObjectFullNames] = GetObjectsVolume(this, ObjectFullNames)
            this.CheckProjectIsOpen(true);
            
            % if object names are NOT provided, get name of all objects in the project 
            if nargin<2 || isempty(ObjectFullNames),
                ObjectFullNames = this.GetAllObjectNames();
            % otherwise, check them
            else
                ObjectFullNames = this.CheckObjectFullNameAndExistance(ObjectFullNames);
            end
            
            Solid = this.FProj.invoke('Solid');
            NObjs = length(ObjectFullNames);
            Vol = nan(NObjs,1);
            for iobj=1:NObjs
                Vol(iobj) = Solid.invoke('GetVolume',ObjectFullNames{iobj});
            end
        end

        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [Mass, ObjectFullNames] = GetObjectsMass(this, ObjectFullNames)
            this.CheckProjectIsOpen(true);

            % if object names are NOT provided, get name of all objects in the project 
            if nargin<2 || isempty(ObjectFullNames),
                ObjectFullNames = this.GetAllObjectNames();
            % otherwise, check them
            else
                ObjectFullNames = this.CheckObjectFullNameAndExistance(ObjectFullNames);
            end
            
            Solid = this.FProj.invoke('Solid');
            NObjs = length(ObjectFullNames);
            Mass = nan(NObjs,1);
            for iobj=1:NObjs
                Mass(iobj) = Solid.invoke('GetMass',ObjectFullNames{iobj});
            end
        end

        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [ObjectMaterials, ObjectFullNames] = GetObjectsMaterial(this, ObjectFullNames)
            this.CheckProjectIsOpen(true);
            
            % if object names are NOT provided, get name of all objects in the project 
            if nargin<2 || isempty(ObjectFullNames),
                ObjectFullNames = this.GetAllObjectNames();
            else % otherwise, check them
                ObjectFullNames = this.CheckObjectFullNameAndExistance(ObjectFullNames);
            end
            
            Solid = this.FProj.invoke('Solid');
            NObjs = length(ObjectFullNames);
            ObjectMaterials = cell(NObjs,1);
            for iobj=1:NObjs
                ObjectMaterials{iobj} = Solid.invoke('GetMaterialNameForShape',ObjectFullNames{iobj});
            end
            
            if NObjs==1, ObjectMaterials = ObjectMaterials{1}; end
        end

        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function MaterialNames = GetMaterialNames(this)
            this.CheckProjectIsOpen(true);

            % Get names of all materials used in the project  
            Material = this.FProj.invoke('Material');
            NMat = Material.invoke('GetNumberOfMaterials');
            MaterialNames  = cell(NMat,1);
            for imat = 1:NMat
                MaterialNames{imat} = Material.invoke('GetNameOfMaterialFromIndex',imat-1);
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [MaterialColors, MaterialNames, MaterialAlphas] = GetMaterialColors(this)
            this.CheckProjectIsOpen(true);
            
            % Define file names
            [~, MacroFile, DataFile] = this.GetTempDirAndFiles('GetMaterialColor');
            
            % Write macro file and execute it
            this.WriteNamedScriptToFile('GetMaterialColor', MacroFile, {'TxtFileName',DataFile});
            this.FProj.invoke('RunScript',MacroFile);
            % Read data saved by the macro
            d = importdata(DataFile);
            MaterialColors = d.data;
            MaterialNames = d.textdata;
            
            % If we want to get the material transparency.
            if nargout>=3,
                % It seems there is no API to get the material transparency... 
                % So, we will do a work-around (again...).
                % The idea is - to parse the model history list for the
                % commands like 
                %    With Material 
                %        .Name "PTFE (lossy)"
                %        .Folder ""
                %        .Colour "0", "0.501961", "1" 
                %        .Wireframe "False" 
                %        .Reflection "False" 
                %        .Allowoutline "True" 
                %        .Transparentoutline "False" 
                %        .Transparency "10" 
                %        .ChangeColour 
                %    End With 
                % and get the materials' Transparency from them.
                % It would be better to get the last history list from the
                % project, but I could not find a way how... Therefore, we will
                % parse the project's .mod file instead. Therefore, the
                % project must have been saved after changing the material's transparency 
                % in order to get its updated value.
                % Anyway, this is an experimental feature...
            
                NMat = length(MaterialNames);
                
                % Alpha is equal to (1-Transparency), where Transparency is between 0 and 1
                % If material will not be found below, consider this material fully opaque (Aplha=1 or Transparency=0). 
                MaterialAlphas = ones(NMat,1); 

                try
                    % Read the model .mod file.
                    Model3DDir = this.FProj.invoke('GetProjectPath', 'Model3D');
                    [~,Txt] = this.ReadTextFileContentToLines(fullfile(Model3DDir,'Model.mod'));
                    
                    % Find all material names and their transparencies  
                    d = regexp(Txt.', 'With Material.*?\.Name\s*"([^"]*)".*?\.Transparency\s*"([^"]*)"', 'Tokens');
                    if isempty(d), return; end
                    NBlocks = length(d);
                    
                    % 
                    for ibl=1:NBlocks
                        assert(length(d{ibl})==2, 'Parse error: Expected length of the token is 2, but we got %i...', length(d{ibl}));
                        ModMatName = d{ibl}{1};
                        ModMatTransp = str2double(d{ibl}{2});
                        assert(~isnan(ModMatTransp), 'Parse error: Material transparency is not a number. Bug?');
                        [~,ia] = intersect(MaterialNames,ModMatName);
                        if ~isempty(ia)
                            MaterialAlphas(ia) = 1 - ModMatTransp/100;
                        end
                    end
                    
                catch ME
                    warning('TCSTInterface:GetMatTranspError', 'Cannot get the material transparensy due to the following error:\n%s',ME.message);
                end
                
                
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [ObjectColors, ObjectFullNames, ObjectAlphas] = GetObjectColors(this, ObjectFullNames)
            this.CheckProjectIsOpen(true);
            GetAlphas = (nargout>=3);
            
            % if object names are NOT provided, we will get name of all objects in the project later 
            if nargin<2,
                ObjectFullNames = [];
            else % otherwise, check them
                ObjectFullNames = this.CheckObjectFullNameAndExistance(ObjectFullNames);
            end
            
            % Get material of the objects, and object names, if they are not provided 
            [ObjectMaterials, ObjectFullNames] = this.GetObjectsMaterial(ObjectFullNames);
            % Get color of all project materials
            if GetAlphas
                [MaterialColors, MaterialNames, MaterialAlphas] = this.GetMaterialColors();
            else
                [MaterialColors, MaterialNames] = this.GetMaterialColors();
            end
            
            % Find all objects' color
            [res,ib] = ismember(ObjectMaterials, MaterialNames);
            assert(all(res), 'One or more object materials were not found! Bug?');
            ObjectColors = MaterialColors(ib,:);
            if GetAlphas
                ObjectAlphas = MaterialAlphas(ib);
            end
            
            % Find objects which have individual colors and replace
            % material colors for these object by the individual ones
            [ind, Colors] = GetObjectIndividualColors(this, ObjectFullNames);
            ObjectColors(ind,:) = Colors(ind,:);
            
            % Helping function: Tries to obtain the objects' individual colors  
            function [iObjsWithOwnColor, Colors] = GetObjectIndividualColors(this, ObjectFullNames)
                % It seems there is no API to get the object's individual
                % color... So, we will do a work-around (again...).
                % The idea is - to parse the model history list for the
                % commands like 
                %    Solid.ChangeIndividualColor "component1:Ground", "0", "140", "0"
                % and get the objects' individual colors from them.
                % It would be better to get the last history list from the
                % project, but I could not find a way how... Therefore, we will
                % parse the project's .mod file instead. Therefore, the
                % project must have been saved after changing the objects
                % color in order to get its updated color.
                % Anyway, this is an experimental feature...

                NObjs = length(ObjectFullNames);
                iObjsWithOwnColor = false(NObjs,1);
                Colors = nan(NObjs,3);

                try
                    % Read the model .mod file.
                    Model3DDir = this.FProj.invoke('GetProjectPath', 'Model3D');
                    Lines = this.ReadTextFileContentToLines(fullfile(Model3DDir,'Model.mod'));
                    % Check each object
                    for iobj=1:NObjs,
                        iLns = find( contains(Lines, ['Solid.ChangeIndividualColor "' ObjectFullNames{iobj} '"']) );
                        if isempty(iLns), continue; end
                        Ln = Lines{iLns(end)}; % there may be several such changes, take the last one
                        d = regexp(Ln,'"(\d+)",\s?"(\d+)",\s?"(\d+)"', 'tokens');
                        assert(length(d{1})==3, 'An RGB triplet is expected, but got %i numeric values.\nThe object name is "%s".', length(d{1}), ObjectFullNames{iobj});
                        iObjsWithOwnColor(iobj) = true;
                        Colors(iobj,:) = cellfun(@(c) str2double(c), d{1}) / 255;
                    end
                catch ME
                    warning('TCSTInterface:GetIndClrsError', 'Cannot get the individual colors for objects due to the following error:\n%s',ME.message);
                end
            end
        end

        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function WireFrame(this, State)
            this.CheckProjectIsOpen(true);
            assert(islogical(State), '"State" must have a logical value.');
            
            PlotObj = this.FProj.invoke('Plot');
            PlotObj.invoke('WireFrame',State);
            PlotObj.invoke('Update');
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function State = GradientBackground(this, State)
            this.CheckProjectIsOpen(true);
            
            PlotObj = this.FProj.invoke('Plot');
            if nargin>=2,
                PlotObj.invoke('SetGradientBackground',State);
                PlotObj.invoke('Update');
            end
            if nargout>=1,
                State = PlotObj.invoke('GetGradientBackground');
            end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function ExportImage(this, FileName, WidthPx, HeightPx)
            % Supported types are bmp, jpeg and png
            this.CheckProjectIsOpen(true);
            [~,~,Ext] = fileparts(FileName);
            assert(any(strcmpi(Ext,{'.bmp','.jpeg','.png'})), 'Only "bmp", "jpeg" and "png" image types are supported.');
            if nargin<3, WidthPx = 800; end
            if nargin<4, HeightPx = 600; end
            
            PlotObj = this.FProj.invoke('Plot');
            PlotObj.invoke('StoreImage', TCSTInterface.GetFullPath(FileName), WidthPx, HeightPx);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function SelectTreeItem(this,TreeItem)
            this.CheckProjectIsOpen(true);
            this.TreeItemExists(TreeItem, true);
            this.FProj.invoke('SelectTreeItem',TreeItem);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function SetView(this, sView)
            this.CheckProjectIsOpen(true);
            
            if iscell(sView)
                for ic=1:length(sView)
                    this.SetView(sView{ic});
                end
                return
            end

            PlotObj = this.FProj.invoke('Plot');
            % Standard CST views
            if ischar(sView)
                ValidViews = {'Left','Right','Front','Back','Top','Bottom','Perspective','Nearest Axis', 'ZoomToStructure'};
                sView = validatestring(sView, ValidViews, 'SetView', 'sView');
                if strcmp(sView, 'ZoomToStructure')
                    PlotObj.invoke('ZoomToStructure');
                else
                    PlotObj.invoke('RestoreView',sView);
                end
            else
                validateattributes(sView,{'numeric'},{'nonempty','size',[1 2]},'','sView');
                sView(sView<0) = sView(sView<0) + 360;
                
                if all(sView==[0 0]),
                    PlotObj.invoke('RestoreView','Top'); % just to avoid "toggling" view by calling next line when sView=[0 0]
                end
                PlotObj.invoke('RestoreView','Bottom');
                
                if sView(1)>0,
                    PlotObj.invoke('RotationAngle',sView(1));
                    PlotObj.invoke('Rotate','left');
                end
                if sView(2)>0,
                    PlotObj.invoke('RotationAngle',sView(2));
                    PlotObj.invoke('Rotate','down');
                end

            end
            
            PlotObj.invoke('Update');
        end
    
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function ExportObjectsToSTL(this, varargin)
            this.CheckProjectIsOpen(true);
            
            STLFileName = varargin{1};

            % Process arguments
            ValidArgs = { ...
                'Objects', ...
                'AppendColorsToObjName', ... % if true, gets objects' color and append it to object names in the STL file
                'ExportUnits', ...
                'NormalTolerance', ... From CST doc: "Normal tolerance is the maximum angle between any two surface normals on a facet. Set this option to control accuracy of the exported model compared to the model in the project."
                'SurfaceTolerance' ... From CST doc: "Surface tolerance is the maximum distance between the facet and the part of the surface it is representing. Set this option to control accuracy of the exported model compared to the model in the project."
            };
            funcName = 'ExportObjectsToSTL';
            ObjectFullNames       = {};
            AppendColorsToObjName = false;
            ExportUnits           = 'mm';
            NormalTolerance       = [];
            SurfaceTolerance      = [];
            iarg = 2;
            while iarg<=length(varargin),
                ArgName = validatestring(varargin{iarg}, ValidArgs, funcName, '', iarg);
                switch ArgName,
                    case 'Objects'
                        this.CheckIfValueSpecified(varargin,iarg, {'char','cell'});
                        ObjectFullNames = this.CheckObjectFullNameAndExistance(varargin{iarg+1});
                        iarg = iarg+2;
                    case 'AppendColorsToObjName',
                        if iarg<length(varargin) && ~ischar(varargin{iarg+1})
                            AppendColorsToObjName = varargin{iarg+1};
                            assert(islogical(AppendColorsToObjName), '"AppendColorsToObjName" must have logical value.')
                            iarg = iarg+2;
                        else
                            AppendColorsToObjName = true;
                            iarg = iarg+1;
                        end
                    case 'ExportUnits'
                        this.CheckIfValueSpecified(varargin,iarg, {'char'});
                        ValidUnits = {'m','cm','mm','um','ft','in','mil'};
                        ExportUnits = validatestring(varargin{iarg+1}, ValidUnits, funcName, ArgName, iarg);
                        iarg = iarg+2;
                    case 'NormalTolerance'
                        this.CheckIfValueSpecified(varargin,iarg, {'numeric'});
                        NormalTolerance = varargin{iarg+1};
                        iarg = iarg+2;
                    case 'SurfaceTolerance'
                        this.CheckIfValueSpecified(varargin,iarg, {'numeric'});
                        SurfaceTolerance = varargin{iarg+1};
                        iarg = iarg+2;
                end
            end
            % if object names are NOT provided, get name of all objects in the project 
            if isempty(ObjectFullNames),
                ObjectFullNames = this.GetAllObjectNames();
            end
            
            % If we want to append colors to the object name, 
            if AppendColorsToObjName,
                [ObjectColors,~,ObjectAlphas] = this.GetObjectColors(ObjectFullNames);
            end
            
            % STL object in CST supports export of only a sigle object... 
            % Therefore, we export all objects one by one to a temp file, 
            % and afterwards combine them in the final file. 
            % We will also replace the solid name (which is just a path to the file) 
            % in the STL file by its object name.
            % Additionally, if the option "AppendColorsToObjName" is true,
            % we get the objects' color and transparency and append them to
            % the object name.
            
            % Open the final file
            fidW = fopen(STLFileName,'w');
            if fidW<0, error('CSTInterface:STLCanNotWrite', 'Can''t open file "%s" for writing.', STLFileName); end
            try
                
                NObjs = length(ObjectFullNames);
                for iobj=1:NObjs
                    d = strsplit(ObjectFullNames{iobj},':');
                    ObjectComponent = d{1};
                    ObjectName      = d{2};

                    % Export the object to temporal STL file
                    TempSTLFile = [tempname(this.GetTempDirAndFiles) '.stl'];
                    % ... write STL
                    STL = this.FProj.invoke('STL');
                    STL.invoke('Reset');
                    STL.invoke('FileName',        TempSTLFile     );
                    STL.invoke('Name',            ObjectName      );
                    STL.invoke('Component',       ObjectComponent );
                    STL.invoke('ScaleToUnit',     true            );
                    STL.invoke('ExportFileUnits', ExportUnits     );
                    STL.invoke('ExportFromActiveCoordinateSystem',false);
                    if ~isempty(NormalTolerance)
                        STL.invoke('NormalTolerance',NormalTolerance);
                    end
                    if ~isempty(SurfaceTolerance)
                        STL.invoke('SurfaceTolerance',SurfaceTolerance);
                    end
                    STL.invoke('Write');

                    % Read content, fix the "solid" name and append to the final file. 
                    % ... read the content and delete the temp file
                    try
                        Lines = this.ReadTextFileContentToLines(TempSTLFile);
                    catch ME
                        if isfile(TempSTLFile), delete(TempSTLFile); end
                        rethrow(ME);
                    end
                    delete(TempSTLFile); % delete the temp file
                    % ... form the solid name
                    SolidName = ObjectFullNames{iobj};
                    if AppendColorsToObjName,
                        c = ObjectColors(iobj,:);
                        SolidName = sprintf('%s COLOR_RGBA={%g,%g,%g,%g}',SolidName,c(1),c(2),c(3),ObjectAlphas(iobj));
                    end
                    % ... replace the corresponding line in the Lines 
                    iLnSt = find(startsWith(Lines,'solid'),1);
                    assert(~isempty(iLnSt), 'CSTInterface:STLNoSolidFound', 'STL file format mismatch: cannot find sting "solid". Is it binary STL???')
                    Lines{iLnSt} = ['solid ' SolidName];
                    % ... replace the solid name at the "endsolid" as well  
                    iLnEnd = find(startsWith(Lines,'endsolid'),1);
                    assert(~isempty(iLnEnd), 'CSTInterface:STLNoEndsolidFound', 'STL file format mismatch: cannot find sting "endsolid". Is it binary STL???')
                    Lines{iLnEnd} = ['endsolid ' SolidName];
                    % ... append the solid data to the final file
                    fprintf(fidW,'%s\n',Lines{iLnSt:iLnEnd});
                end
                
            catch ME
                fclose(fidW);
                rethrow(ME);
            end
            fclose(fidW);
            
        end

        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function varargout = DisplayTreeItems(this, RootItem)
            
            this.CheckProjectIsOpen(true);
            if (nargin<2),  RootItem = [];  end
            
            [TreeItems, Levels] = this.EnumerateTreeItems(RootItem);

            fprintf('\n');
            n = fprintf('<strong>-------------- Result tree under "%s" ---------------</strong>\n', RootItem);
            for k=1:length(TreeItems)
                fprintf('%s%s\n', repmat(' | ',1,Levels(k)), TreeItems{k});
            end
            fprintf('<strong>%s</strong>\n\n',repmat('-',1,n-18));
            
            % we use varargout (not directly output arguments) in order to prevent printing the TreeItems cell array if this method was called in command window without semicolon and without output args.  
            if nargout>=1,  varargout{1} = TreeItems;  end
            if nargout>=2,  varargout{2} = Levels;     end
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [TreeItems, Levels] = EnumerateTreeItems(this, varargin)
            this.CheckProjectIsOpen(true);
            
            if nargin<2 || isempty(varargin{1}), RootItem = '1D Results';
            else, RootItem = varargin{1};
            end
            this.TreeItemExists(RootItem, true);
            
            ValidArgs = { ...
                'DoesNotHaveChildren' ...
            };
            funcName = 'EnumerateTreeItems';
            iarg = 2;
            DoesNotHaveChildren = false;
            while iarg<=length(varargin),
                ArgName = validatestring(varargin{iarg}, ValidArgs, funcName, '', iarg);
                switch ArgName,
                    case 'DoesNotHaveChildren',
                        this.CheckIfValueSpecified(varargin,iarg, {'logical'});
                        DoesNotHaveChildren = varargin{iarg+1};
                        iarg = iarg+2;  
                end
            end

            TreeItems = cell(0,1);
            Levels = nan(500,1);
            ResultTree = this.FProj.invoke('Resulttree');
            [TreeItems, Levels] = localEnumTreeItemsRecursively(ResultTree, RootItem, 0, TreeItems, Levels);
            Levels(isnan(Levels)) = [];
            
            if DoesNotHaveChildren
                NItems = length(TreeItems);
                iItemsToRemove = false(NItems,1);
                for n=1:NItems-1
                    iItemsToRemove(n) = Levels(n+1)>Levels(n);
                end
                TreeItems(iItemsToRemove) = [];
                Levels(iItemsToRemove) = [];
            end

            function [TreeItems, Levels] = localEnumTreeItemsRecursively(ResultTree, RootItem, Level, TreeItems, Levels)
                if isempty(RootItem), return; end
                TreeItem = ResultTree.invoke('GetFirstChildName',RootItem);
                while ~isempty(TreeItem),
                    TreeItems{end+1,1} = TreeItem; %#ok<AGROW>
                    Levels(length(TreeItems)) = Level;
                    [TreeItems, Levels] = localEnumTreeItemsRecursively(ResultTree, TreeItem, Level+1, TreeItems, Levels);
                    TreeItem = ResultTree.invoke('GetNextItemName',TreeItem);
                end
            end
        end
        
        
        % -----------------------------------------------------------------
        % Gets/prints the license HostID and Customer number. It may be 
        % useful e.g. to create an CST account for the support.
        % -----------------------------------------------------------------
        function [HostID, CustNum] = PrintLicenseInfo(this)
            % a project must be openned to use this function 
            this.ConnectToCSTOrStartIt();
            if ~this.CheckProjectIsOpen(false),
                % try to get currently open project
                this.OpenProject();
            end
            
            HostID  = this.FProj.invoke('GetLicenseHostId');
            CustNum = this.FProj.invoke('GetLicenseCustomerNumber');
            
            if nargout<2,
                fprintf('-------- License info ---------\n');
                fprintf('<strong>License Host Id:</strong> %s\n', HostID);
                fprintf('<strong>License Customer Number:</strong> %s\n', CustNum);
                fprintf('-------------------------------\n');
            end

        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function SetMatlabCostFunction(this, CostFunction)
            this.CheckProjectIsOpen(true);
            
            % Check function exist and get its path and name
            assert(exist(CostFunction,'file')>0, 'Function "%s" does not exist in Matlab path.', CostFunction);
            [CostFnPath,CostFnName] = fileparts( TCSTInterface.GetFullPath(CostFunction) );
            
%             % Store parameters with the function path and name IN THE DESCRIPTION of the parameters. They will be used in VBA script in CST.
%             % Value of these parameters is not important and it is ignored.
%             this.StoreParameterWithDescription('MLCostFnPath', 0, CostFnPath, false); % "false" since no structure update is required for that.
%             this.StoreParameterWithDescription('MLCostFnName', 0, CostFnName, false);
            
            % Write VBA script to the CST-project directory
            ScriptName = 'CalcCostFnInMatlab.mcr';
            Model3DDir = this.FProj.invoke('GetProjectPath', 'Project');
%             this.WriteNamedScriptToFile('ExecMatlabInTBPP', fullfile(Model3DDir,ScriptName), {});
            this.WriteNamedScriptToFile('ExecMatlabInTBPP', fullfile(Model3DDir,ScriptName), {'MLCostFnPath',CostFnPath; 'MLCostFnName',CostFnName});
            
            % Enable COM server in Matlab
            enableservice('AutomationServer',true);
            
            % Print what has to be done manually in CST project.
            % Unfortunatelly, I could not find a way to add a Template Based Post-Processing step programmatically... 
            this.fprintf('\n-----------------------------------------------------------------------------------------------------------------\n');
            this.fprintf('<strong>To finalize optimization setup in CST do the following steps:</strong>\n');
            this.fprintf('  1. Add Template Based Post-Processing step:\n');
            this.fprintf('    a) Go to Template Based Post-Processing (Shift+P) \n');
            this.fprintf('    b) Under category "Misc" select "Run VBA Code"\n');
            this.fprintf('    c) Place the following code there and press Ok:\n');
            this.fprintf('<strong>Sub Main</strong>\n');
            this.fprintf('<strong>  RunScript(GetProjectPath("Project")+GetPathSeparator()+"CalcCostFnInMatlab.mcr")</strong>\n');
            this.fprintf('<strong>End Sub</strong>\n');
            this.fprintf('    d) (optionally) Select added post-processing step and press "Evaluate" to check that Matlab script is called and works as expected\n');
            this.fprintf('  2. Add optimization goal:\n');
            this.fprintf('    a) Go to Simulation tab - Optimizer - Goals - Add New Goal\n');
            this.fprintf('    b) Under "Result Name" select newly added post-Processing step (like "TBPP 0D: Run VBA Code")\n');
            this.fprintf('    c) Complete setting the goal and press Ok\n');
            this.fprintf('Now the project is ready to evaluate the cost function in Matlab.\n');
            this.fprintf('-----------------------------------------------------------------------------------------------------------------\n\n');
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function About(this)
            c = clock;
            this.fprintf('\n');
            this.fprintf('*****************************************************************************\n');
            this.fprintf('***               <strong>CST interface v%s  for  CST %s</strong>                   ***\n', this.FVersion, this.FCSTVersion);
            this.fprintf('***             %s                 ***\n', this.FOrganization);
            this.fprintf('***      %s         ***\n', this.FAuthor);
            this.fprintf('***                     Last edit:  %s                            ***\n', this.FLastEdit);
            this.fprintf('***                       Today: %02i-%02i-%i                               ***\n', c(3), c(2), c(1));
            this.fprintf('*****************************************************************************\n\n');
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [Units, Quantity, ConversionCoefToSI] = GetProjectUnits(this, Quantity)
            this.CheckProjectIsOpen(true);
            
            % Columns: Quantity name, API to get units, API to get the conversion coefficient 
            QuantitiesAndAPIs = {
                'Length',       'GetGeometryUnit',      'GetGeometryUnitToSI'; ...
                'Time',         'GetTimeUnit',          'GetTimeUnitToSI'; ...
                'Frequency',    'GetFrequencyUnit',     'GetFrequencyUnitToSI'; ...
                'Voltage',      'GetVoltageUnit',       'GetVoltageUnitToSI'; ...
                'Current',      'GetCurrentUnit',       'GetCurrentUnitToSI'; ...
                'Resistance',   'GetResistanceUnit',    'GetResistanceUnitToSI'; ...
                'Conductance',  'GetConductanceUnit',   'GetConductanceUnitToSI'; ...
                'Capacitance',  'GetCapacitanceUnit',   'GetCapacitanceUnitToSI'; ...
                'Inductance',   'GetInductanceUnit',    'GetInductanceUnitToSI'; ...
                'Temperature'   'GetTemperatureUnit',   'GetTemperatureUnitToSI' ...
            };
            
            % If Quantity in not provided, return units for all of them
            if nargin<2 || isempty(Quantity)
                NQ = size(QuantitiesAndAPIs,1);
                Units = cell(NQ,1);
                ConversionCoefToSI = nan(NQ,1);
                for iq=1:NQ,
                    [Units{iq},~,ConversionCoefToSI(iq)] = this.GetProjectUnits(QuantitiesAndAPIs{iq,1});
                end
                Quantity = QuantitiesAndAPIs(:,1);
                return
            end
            
            % If Quantity is a cell array, run this method for each its cell 
            if iscell(Quantity)
                NQ = numel(Quantity);
                QuantityIn = Quantity;
                Quantity = cell(NQ,1); % we will return Quantity passed throung this method, so it will be exactly like in QuantitiesAndAPIs(:,1)   
                Units = cell(NQ,1);
                ConversionCoefToSI = nan(NQ,1);
                for iq=1:NQ,
                    [Units{iq},Quantity{iq},ConversionCoefToSI(iq)] = this.GetProjectUnits(QuantityIn{iq});
                end
                return
            end
        
            % Validate the quantity name and find its index
            Quantity = validatestring(Quantity, QuantitiesAndAPIs(:,1), 'GetProjectUnits', 'Quantity');
            ind = find( strcmp(QuantitiesAndAPIs(:,1),Quantity), 1);
            
            % Call CST's APIs
            UnitsObj = this.FProj.invoke('Units');
            Units              = UnitsObj.invoke(QuantitiesAndAPIs{ind,2});
            ConversionCoefToSI = UnitsObj.invoke(QuantitiesAndAPIs{ind,3});
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [ParamNames, ParamValues, ParamMinMax, ParamInitValues] = GetOptimizerVaryingParameters(this)
            this.CheckProjectIsOpen(true);
            
            OptObj = this.FProj.invoke('Optimizer');
            NParams = OptObj.invoke('GetNumberOfVaryingParameters');
            ParamNames      = cell(NParams,1);
            ParamInitValues = nan(NParams,1);
            ParamMinMax     = nan(NParams,2);
            ParamValues     = nan(NParams,1);
            for ip=1:NParams
                ParamNames{ip}      = OptObj.invoke('GetNameOfVaryingParameter',ip-1);
                ParamInitValues(ip) = OptObj.invoke('GetParameterInitOfVaryingParameter',ip-1);
                ParamMinMax(ip,1)   = OptObj.invoke('GetParameterMinOfVaryingParameter',ip-1);
                ParamMinMax(ip,2)   = OptObj.invoke('GetParameterMaxOfVaryingParameter',ip-1);
                ParamValues(ip)     = OptObj.invoke('GetValueOfVaryingParameter',ip-1);
            end
            
        end
%         
%         % -----------------------------------------------------------------
%         % 
%         % -----------------------------------------------------------------
%         function (this)
%             
%         end
        
%         % -----------------------------------------------------------------
%         % 
%         % -----------------------------------------------------------------
%         function [S, varargout] = GetSParamsAsFunctionOf(this, Arguments, varargin)
%         % [S, ArgumentsValues{:}] = GetSParamsAsFunctionOf(this, Arguments, [ConstParam1Name,ConstParam1Value, ...]);
%             
%             error('TO DO');
%         
%             % make sure that a project is open in CST 
%             this.CheckProjectIsOpen(true);        
%             
%             % Check arguments
%             assert(ischar(Arguments)||iscell(Arguments), '"Arguments" must be a string or a cell array of strings with parameter names or "Freq".')
%             if ischar(Arguments),
%                 Arguments = {Arguments};
%             end
%             NArgs = length(Arguments);
%             for iarg=1:NArgs,
%                 assert(ischar(Arguments{iarg}), 'All elements of the cell array "Arguments" must be parameter names or "Freq".');
%             end
%             
%             % get constant params
%             assert(mod(length(varargin),2)==0, 'Constant parameters must be specified as (Name,Value) pairs.');
%             NConstParams = length(varargin)/2;
%             ConstParamNames  = cell(NConstParams,1);
%             ConstParamValues = nan(NConstParams,1);
%             for ipar=1:NConstParams,
%                 ConstParamNames{ipar}  = varargin{2*(ipar-1)+1};
%                 ConstParamValues(ipar) = varargin{2*ipar};
%             end
%             
%             % Get avalable parameter names
%             ParamNames = this.GetParameterCombination();
%             
%             % if there is a parameter with name "Freq". 
%             % 
%             
%             % Check that specified Arguments and const params exist
%             for iarg=1:length(ParamNames),
%                 assert(ismember(Arguments{iarg},ParamNames) || strcmp(Arguments{iarg},'Freq'))
%             end
%         
%         end

    end
end





%% VB scripts for CST 
% !!! WARNING !!! VARNING !!! WARNUNG !!! AVVERTENZA !!! AVERTISSEMENT !!!
% The content of this cell is VBA scripts which are written to files by the
% private method WriteNamedScriptToFile.
% !!! Nevertheless everything below is commented, DO NOT CHANGE this cell 
% unless you know what you are doing !!! 

% =========================================================================
%                              SCRIPT 1
% This script is used in GetParameterCombination method to work-around
% inability to execute the CST function GetParameterCombination directly.
% TO DO: Add option to get param combinations for ALL run IDs (often used) 
% =========================================================================
%>>> Name: GetParameterCombination
% Sub Main ()
%   Dim names As Variant, values As Variant, exists As Boolean
%   exists = GetParameterCombination( "${ResID}", names, values )
%   Open "${DataFile}" For Output As #1
%   If Not exists Then
%     Print #1,  "Parameter combination does not exist."
% '    ReportInformationToWindow( "Parameter combination does not exist."
%   Else
%     Dim N As Long
%     For N = 0 To UBound( values )
% '      ReportInformationToWindow( names( N )  + ": " + CStr( values( N ) ) )
%       Print #1,  names( N ) + ", " + CStr( values( N ) )
%     Next
%   End If
%   Close #1
% End Sub
%<<< END OF SCRIPT
% =========================================================================

% =========================================================================
%                               SCRIPT 2
% This script runs Matlab function and stores the returned real scalar as
% returned value of the Template Based Postprocessing entry "Run VBA script". 
% Used for optimization in CST with the cost function evaluated in Matlab.
% =========================================================================
%>>> Name: ExecMatlabInTBPP
% Dim ML As Object
% 
% Sub Main () 
% 
% 	Dim CostFnName As String, CostFnPath As String
% 	Dim tiCostFn As Object
% 	Dim CostFnVal As Double
% 	Dim x As Double
% 
% 	' CostFnPath = GetParameterDescription("MLCostFnPath")
% 	' CostFnName = GetParameterDescription("MLCostFnName")
% 	CostFnPath = "${MLCostFnPath}"
% 	CostFnName = "${MLCostFnName}"
% 
% 	ReportInformationToWindow("------------------------------>")
% 
% 	' Connect to the Matlab COM server
% 	ReportInformationToWindow("Connecting to Matlab...")
% 	On Error GoTo MLNotOpened
% 	Set ML = GetObject(, "Matlab.Desktop.Application")
% 	On Error GoTo 0
% 
% 	' Evaluate Matlab function CostFnName
% 	ReportInformationToWindow("Evaluating '"+CostFnName+"' in Matlab...")
% 	MLEval("clear; clc;")
% 	MLEval("cd('" + CostFnPath + "');")
% 	MLEval("CostFnVal = " + CostFnName + "('" + GetProjectPath("Project") + "');")
% 	CostFnVal = MLGetRealScalar("CostFnVal")
% 	ReportInformationToWindow("CostFn value is: " + CStr(CostFnVal))
% 
% 	StoreGlobalDataValue("EVALUATE0DRETURN", CostFnVal) ' return a value to the PP template
% 
% 	ReportInformationToWindow("<------------------------------")
% 	Exit Sub
% 
% MLNotOpened:
% 	'ReportError ("Please start MATLAB with '-automation' key")
% 	ReportError ("Please start MATLAB and execute ""enableservice('AutomationServer',true);"" there.")
% 
% End Sub
% 
% ' --------------------------------------------------------------------
% ' Evaluates Matlab command and checks if there was an error
% ' --------------------------------------------------------------------
% Sub MLEval(EvalStr As String)
% 	Dim Result As String
% 
% 	Result = ML.Execute(EvalStr)
% 	'If InStr(Result, "Error ")>0 Then
% 	If InStr(Result, "???")=1 Then
% 		ReportError("Error in Matlab function:" + vbNewLine + Result)
% 	End If
% End Sub
% 
% ' --------------------------------------------------------------------
% ' Returns specified variable from the ML base workspace as real scalar
% ' --------------------------------------------------------------------
% Function MLGetRealScalar(VarName As String) As Double
% 	Dim Result As String
% 	Dim VarMatRe(0,0) As Double, VarMatIm(0,0) As Double
% 
% 	On Error GoTo GetMatErr
% 	ML.GetFullMatrix(VarName,"base",VarMatRe,VarMatIm)
% 	On Error GoTo 0
% 	MLGetRealScalar = VarMatRe(0,0)
% 
% 	Exit Function
% 
% GetMatErr:
% 	ReportError("Cannot get value of variable '"+VarName+"' from Matlab base workspace. Does it exist there?")
% 
% End Function
%<<< END OF SCRIPT
% =========================================================================

% =========================================================================
%                               SCRIPT 3
% Similarly to the SCRIPT 1, this script is a wrapper for the
% Material.GetColour CST API.
% =========================================================================
%>>> Name: GetMaterialColor
% Sub Main
% 
% 	Dim R As Double, G As Double, B As Double
% 	Dim MatName As String
% 
% 	Open "${TxtFileName}" For Output As #1
% 
% 	For iMat = 0 To Material.GetNumberOfMaterials-1
% 		MatName = Material.GetNameOfMaterialFromIndex(iMat)
% 		Material.GetColour(MatName,R,G,B)
% 		ReportInformationToWindow(MatName+": "+CStr(R)+", "+CStr(G)+", "+CStr(B))
% 		Print #1, MatName + ", " + CStr(R) + ", " + CStr(G) + ", " + CStr(B)
% 	Next
% 
% 	Close #1
% 
% End Sub
%<<< END OF SCRIPT






