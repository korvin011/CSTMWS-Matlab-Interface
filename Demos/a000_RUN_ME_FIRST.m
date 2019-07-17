% set currect directory
cd(fileparts(mfilename('fullpath')));
% add path to helping functions
addpath('demoHelpingFunctions');
addpath('..\Functions');

% We need to do it in m-file, because (rediculus, but true) there is no way
% to obtain current directory within a Live Script. At least I haven't found
% a way in Matlab 2018b...