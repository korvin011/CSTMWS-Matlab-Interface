classdef TSTLReader < handle
    
% @TITLE ~-- 
% 
% @DESCRIPTION
%   Description
% 
% @AUTHOR    Oleg Iupikov <oleg.iupikov@chalmers.se, lichne@gmail.com>

    
    properties (Constant, Access = private)
%         SETTINGS_FILE = fullfile(prefdir,'STLReader.mat');
        FVersion = '1.1.007'  % version
        FAuthor = 'Oleg Iupikov, <a href="mailto:lichne@gmail.com">lichne@gmail.com</a>, <a href="mailto:oleg.iupikov@chalmers.se">oleg.iupikov@chalmers.se</a>';
        FOrganization = 'Chalmers University of Technology, Sweden';
        FFirstEdit = '25-02-2019';
        FLastEdit  = '16-07-2019';
    end
    
    properties
        SolidNames
        Vertices
        Faces
        SolidColors
        SolidAlphas
    end
    
%     properties (Access = private)
%         
%     end
%     
%     properties (Dependent)
%         
%     end
%     
%     % =====================================================================
%     % ==================== GET and SET methods ============================
%     % =====================================================================
%     methods
% %         % -----------------------------------------------------------------
% %         % 
% %         % -----------------------------------------------------------------
% %         function data = get.(this)
% %             data = this.F;
% %         end
% %         function set.(this, data)
% %             this.F = data;
% %         end
%     end
    
    
        % @DESCRIPTION
        %   
        % 
        % @SYNTAX
        % 
        % 
        % 
        % @EXAMPLES
        % 
        % 
    
    % =====================================================================
    % ======================== STATIC Methods ============================
    % =====================================================================
    methods (Static)
        
        function [Vertices, Faces, SolidNames, Colors, Alphas] = ReadAsciiSTL(FileName)
            
            % Read the file content
            [fid, message] = fopen(FileName);
            if fid<0, error('Can''t open file. System message: ''%s''', message); end
            Content = textscan(fid,'%s','delimiter','\n'); % content in a cell array
            Content = Content{1}(~strcmp(Content{:},'')); % remove blank lines
            fclose(fid);
            
            % Name of all solids in the file
            SolidNames = Content(strncmp(Content,'solid',5));
            SolidNames = cellfun(@(c)c(7:end),SolidNames, 'UniformOutput',false);
            
            % Read all solids
            NSolids  = length(SolidNames);
            Vertices = cell(NSolids,1);
            Faces    = cell(NSolids,1);
            Colors   = []; % by default indicate that there is no color info in STL file
            Alphas   = [];
            for isd=1:NSolids,
                SolidName = SolidNames{isd};
                
                % Get content for the current solid
                iLineStart = find(strcmp(Content,['solid ' SolidName]), 1);
                assert(~isempty(iLineStart), 'Could not find beginning of the solid "%s". Bug?',SolidName);
                iLineEnd = find(strcmp(Content,['endsolid ' SolidName]), 1);
                assert(~isempty(iLineEnd), 'Could not find end of the solid "%s". Bug?',SolidName);
                SolidContent = Content(iLineStart+1:iLineEnd-1);
                
                % read the vertex coordinates (vertices)
                strVert = char(SolidContent(strncmp(SolidContent,'vertex',6)));
                SolidVertices = str2num(strVert(:,7:end)); %#ok<ST2NM>
                NVerts = size(strVert,1); % number of vertices
                NFaces = sum(strcmp(SolidContent,'endfacet')); % number of faces
                if (NVerts == 3*NFaces)
                    SolidFaces = reshape(1:NVerts,[3 NFaces])'; % create faces
                else
                    SolidFaces = {};
                    warning('TSTLReader:NonTriFacets', 'Some of the facets in the solid "%s" are not triangular. Face creation skipped.', SolidName);
                end
                
                % Remove idendical vertices
                [SolidVertices, ~, ind] =  unique(SolidVertices, 'rows');
                if ~isempty(SolidFaces),
                    SolidFaces = ind(SolidFaces);
                end
                
                % Check if the solid color is appended
                NumPat = '(?:[-+]?\d*\.?\d+)(?:[eE](?:[-+]?\d+))?';
                [d, ist] = regexp(SolidName, ['\sCOLOR_RGBA={(' NumPat '),(' NumPat '),(' NumPat '),(' NumPat ')}'], 'tokens', 'start');
                if ~isempty(d) % color info found
                    if isempty(Colors)
                        Colors = nan(NSolids,3);
                        Alphas = nan(NSolids,1);
                    end
                    % convert to numeric RGB triplet
                    assert(length(d{1})==4, 'Hm... RGB triplet and transparency are expected (4 values), but got %i numeric values.\nSolid name is "%s".', length(d{1}), SolidName);
                    RGBA = cellfun(@(c) str2double(c), d{1});
                    Colors(isd,:) = RGBA(1:3); %#ok<AGROW>
                    Alphas(isd)   = RGBA(4);   %#ok<AGROW>
                    % remove color info from the solid name
                    SolidNames{isd} = SolidName(1:ist-1);
                end
                
                % Store
                Vertices{isd}   = SolidVertices;
                Faces{isd}      = SolidFaces;
            end
            
        end
        
    end
    
%     % =====================================================================
%     % ====================== PRIVATE methods ==============================
%     % =====================================================================
    methods (Access = private)
%     methods % TEMP!

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

            if nargin<5, ParAttributes = {}; end
            if nargin<6, FuncName = ''; end


            % if ParClasses provided, check ParName
            ParVal = varargs{iarg+1};
            validateattributes(ParVal,ParClasses,ParAttributes,FuncName,ParName,iarg);
        end

        function iSolids = GetSolidIndex(this, SolidNameOrIndex, ExactMatch)
            if nargin<3,  ExactMatch = false;  end
            if ischar(SolidNameOrIndex),  SolidNameOrIndex = {SolidNameOrIndex};  end
            
            % if SolidNameOrIndex is name(s) of the solid(s), resolve it to index 
            if iscell(SolidNameOrIndex),
                % exact name matching
                if ExactMatch,
                    NSolids = length(SolidNameOrIndex);
                    for isd=1:NSolids,
                        assert(any(strcmpi(this.SolidNames,SolidNameOrIndex{isd})), ...
                            'No solid with name "%s" is found.', SolidNameOrIndex{isd});
                    end
                    [~,~,iSolids] = intersect(SolidNameOrIndex, this.SolidNames); %  <<< One line, but without error handling, so we do it in loop...  
                % solid name contains the specified sub-string
                else
                    iSolids = find( contains(this.SolidNames, SolidNameOrIndex) );
                end
            % if SolidNameOrIndex is already an index of the solid(s)
            else
                assert(isnumeric(SolidNameOrIndex), '"SolidNameOrIndex" must be an index or name of the solid(s).');
                iSolids = SolidNameOrIndex;
                assert(max(iSolids)<=length(this.SolidNames), 'Index of a solid exceeds the number of solids.');
            end
            
        end
        
    end
    
    % =====================================================================
    % ======================= PUBLIC methods ==============================
    % =====================================================================
    methods
        % -----------------------------------------------------------------
        % CONSTRUCTOR
        % -----------------------------------------------------------------
        function this = TSTLReader(varargin)
            narginchk(0,1);
            if nargin>=1, 
                this.ReadSTL(varargin{1});
            end
        end
        
        % -----------------------------------------------------------------
        % TODO: Implement reading both ASCII and binary STLs
        % -----------------------------------------------------------------
        function ReadSTL(this, FileName)
            [this.Vertices, this.Faces, this.SolidNames, this.SolidColors, this.SolidAlphas] = TSTLReader.ReadAsciiSTL(FileName);
        end
        
        % -----------------------------------------------------------------
        % 
        % -----------------------------------------------------------------
        function [hPatches, hFELines] = PlotSolid(this, varargin)
            assert(~isempty(this.Vertices), '"Vertices" must not be empty.');
            assert(~isempty(this.Faces), '"Faces" must not be empty.');
            
            % ----------- Process arguments --------------
            
            if nargin<2 || isempty(varargin{1}),
                NSolidsTotal = length(this.Vertices);
                SolidNameOrIndex = 1:NSolidsTotal;
            else
                SolidNameOrIndex = varargin{1};
            end
            
            ValidArgs = { ...
                'Scale', ... 
                'ScaleOrigin', ...
                'Shift', ...
                'CoorSys', ...
                'CombineEges', ...
                'WireFrame', ...
                'PatchOptions', ...
                'LineOptions', ...
                'PlotFE','FeatureEdges','PlotFeatureEdges', ... % true or false
                'FeatureEdgesAlpha', 'FEAlpha', ... = transparency
                'FeatureEdgesAngleDeg', 'FEAngleDeg', ... filter angle in [deg], specified as a scalar in the range [0,180]. featureEdges returns adjacent triangles that have a dihedral angle that deviates from 180 deg by an angle greater than FeatureEdgesAngleDeg. 
                'FeatureEdgesColor', 'FEColor' ...
            };
            Scale = 1;
            ScaleOrigin = [0 0 0];
            Shift = [0 0 0];
            CoorSys = [];
            CombineEges = false;
            WireFrame = false;
            PatchOptions = {};
            LineOptions = {};
            PlotFE = true;
            FEAlpha = 0.5;
            FEColor = [0 0 0];
            FEAngleDeg = 30;
            iarg = 2;
            while iarg<=length(varargin)
                ArgName = validatestring(varargin{iarg}, ValidArgs, 'PlotSolid', '', iarg);
                switch ArgName
                    case 'Scale'
                        this.CheckIfValueSpecified(varargin,iarg, {'numeric'},{'nonempty','scalar','positive'});
                        Scale = varargin{iarg+1};
                        iarg = iarg+2;                    
                    case 'ScaleOrigin'
                        this.CheckIfValueSpecified(varargin,iarg, {'numeric'},{'nonempty','size',[1 3]});
                        ScaleOrigin = varargin{iarg+1};
                        iarg = iarg+2; 
                    case 'Shift'
                        this.CheckIfValueSpecified(varargin,iarg, {'numeric'},{'nonempty','size',[1 3]});
                        Shift = varargin{iarg+1};
                        iarg = iarg+2;
                    case 'CoorSys'
                        this.CheckIfValueSpecified(varargin,iarg);
                        CoorSys = varargin{iarg+1};
                        if ~isempty(CoorSys),  
                            validateattributes(CoorSys,{'TCoorSys'},{'scalar'},'','CoorSys');
                            CSGlobal = TCoorSys(); 
                        end
                        iarg = iarg+2; 
                    case 'CombineEges'
                        if iarg<length(varargin) && ~ischar(varargin{iarg+1})
                            CombineEges = varargin{iarg+1};
                            validateattributes(CombineEges,{'logical'},{'scalar'},'','CombineEges');
                            iarg = iarg+2;
                        else
                            CombineEges = true;
                            iarg = iarg+1;
                        end
                    case 'WireFrame'
                        if iarg<length(varargin) && ~ischar(varargin{iarg+1})
                            WireFrame = varargin{iarg+1};
                            validateattributes(WireFrame,{'logical'},{'scalar'},'','WireFrame');
                            iarg = iarg+2;
                        else
                            WireFrame = true;
                            iarg = iarg+1;
                        end
                    case {'PlotFE','FeatureEdges','PlotFeatureEdges'}
                        this.CheckIfValueSpecified(varargin,iarg, {'logical'},{'nonempty','scalar'});
                        PlotFE = varargin{iarg+1};
                        iarg = iarg+2; 
                    case 'PatchOptions'
                        this.CheckIfValueSpecified(varargin,iarg, {'cell'},{'nonempty'});
                        PatchOptions = varargin{iarg+1};
                        iarg = iarg+2; 
                    case 'LineOptions'
                        this.CheckIfValueSpecified(varargin,iarg, {'cell'},{'nonempty'});
                        LineOptions = varargin{iarg+1};
                        iarg = iarg+2; 
                    case {'FeatureEdgesAlpha', 'FEAlpha'}
                        this.CheckIfValueSpecified(varargin,iarg, {'numeric'},{'nonempty','scalar','>=',0,'<=',1});
                        FEAlpha = varargin{iarg+1};
                        iarg = iarg+2;
                    case {'FeatureEdgesAngleDeg', 'FEAngleDeg'}
                        this.CheckIfValueSpecified(varargin,iarg, {'numeric'},{'nonempty','scalar','>',0,'<',180});
                        FEAngleDeg = varargin{iarg+1};
                        iarg = iarg+2; 
                    case {'FeatureEdgesColor', 'FEColor'}
                        this.CheckIfValueSpecified(varargin,iarg, {'numeric'},{'nonempty','size',[1 3],'>=',0,'<=',1});
                        FEColor = varargin{iarg+1};
                        iarg = iarg+2; 
                end
    
            end
            
            if WireFrame,
                FEDefaultLineWidth = 2;
                FEAlpha = 1;
            else
                FEDefaultLineWidth = 1;
            end
            
            % --------------------------------------------
            
            iSolids = this.GetSolidIndex(SolidNameOrIndex);
            NSolids = length(iSolids); % number of selected solids to plot

            if NSolids>1 || PlotFE || WireFrame
                hold on
            end
            
            % If colors were is STL file, use them
            if ~isempty(this.SolidColors),
                Colors = this.SolidColors(iSolids,:);
                Alphas = this.SolidAlphas(iSolids);
                % if color for some objects are missing, use gray color for them 
                ind = any(isnan(Colors.'));
                if any(ind)
                    Colors(ind,:) = [1 1 1]*0.7;
                    Alphas(ind) = 1;
                end
            % if there is no color info, just make some
            else
                Colors = lines(NSolids);
                Alphas = ones(NSolids,1);
            end
            
            % --------- Perform transformations ----------
            Verts = cell(NSolids,1);
            for isd=1:NSolids
                V = this.Vertices{iSolids(isd)};
                % Apply scale
                V = Scale*(V-ScaleOrigin) + ScaleOrigin;
                % Apply shift
                V = V + Shift;
                % Apply coordinate system transformation, if CoorSys is specified 
                if ~isempty(CoorSys)
                    V = CoorSys.TransformPointToCS(CSGlobal, V);
                end
                Verts{isd} = V;
            end
            % --------------------------------------------

            % ----------------- Faces --------------------
            % Plot objects' faces if it is not a wireframe plot
            hPatches = [];
            if ~WireFrame
                
                hPatches  = gobjects(NSolids,1);
                for isd=1:NSolids
                    iSolid = iSolids(isd);
                    Faces = this.Faces{iSolid};
                    hPatches(isd) = patch('Vertices',Verts{isd}, 'Faces',Faces, 'FaceAlpha',Alphas(isd), 'EdgeAlpha',0, ...
                            'Tag',this.SolidNames{iSolid}, 'FaceColor',Colors(isd,:), PatchOptions{:}); 
                end
                
            end
            % --------------------------------------------
            
            
            % -------------- Feature edges ---------------
            hFELines = [];
            if PlotFE || WireFrame
                
                hFELines  = gobjects(NSolids,1);
                if CombineEges
                    xe = nan(2,0);  ye = xe;  ze = xe;
                end
                for isd=1:NSolids
                    iSolid = iSolids(isd);
                    Faces = this.Faces{iSolid};
                    tri = triangulation(Faces, Verts{isd});
                    FE = featureEdges(tri, FEAngleDeg/180*pi).';
                    % If there are no FE for this object, there is nothing to draw  
                    if ~isempty(FE)
                        x = tri.Points(:,1);
                        y = tri.Points(:,2);
                        z = tri.Points(:,3);
                        % Combine all feature edges to a single line. This will drastically improve FPS when rotating the object(s)  
                        if CombineEges,
                            xe = [xe x(FE)]; %#ok<AGROW>  Ok to grow, since we don't know in advance how many FE all objects have   
                            ye = [ye y(FE)]; %#ok<AGROW>  Anyway, there should not be too many different objects...
                            ze = [ze z(FE)]; %#ok<AGROW>
                        else
                            xe = [x(FE); nan(1,size(FE,2))];
                            ye = [y(FE); nan(1,size(FE,2))];
                            ze = [z(FE); nan(1,size(FE,2))];
                            if WireFrame
                                FEColor = Colors(isd,:);
                            end
                            hFELines(isd) = plot3(xe(:),ye(:),ze(:),'LineWidth',FEDefaultLineWidth, ...
                                'Tag',this.SolidNames{iSolid}, 'color',[FEColor FEAlpha], LineOptions{:});
                        end
                    end
                end
                if CombineEges,
                    xe = [xe; nan(1,size(xe,2))];
                    ye = [ye; nan(1,size(ye,2))];
                    ze = [ze; nan(1,size(ze,2))];
                    hFELines = plot3(xe(:),ye(:),ze(:),'LineWidth',FEDefaultLineWidth, ...
                        'Tag','FeatureEdges', 'color',[FEColor FEAlpha], LineOptions{:});
                end
                
            end
            % --------------------------------------------
            
            return
            
        end
%         
%         % -----------------------------------------------------------------
%         % 
%         % -----------------------------------------------------------------
%         function 
%             
%         end
%         
%         % -----------------------------------------------------------------
%         % 
%         % -----------------------------------------------------------------
%         function 
%             
%         end
        
    end
end





