%% ========================================================================
% 数据清洗：百度招聘数据（终极修复版）
% 功能：读取原始Excel → 完美修复所有类型错位 → 清洗字段 → 保存
% =========================================================================
clc; clear; close all;

% =============================================
% 0. 直接指定文件完整路径
% =============================================
inputFile = 'D:\Desktop\2026春季《数据挖掘》考试规定动作 + 基础数据.xlsx';
outputExcel = 'D:\Desktop\招聘数据_清洗后.xlsx';
outputCSV = 'D:\Desktop\招聘数据_清洗后.csv';

if ~exist(inputFile, 'file')
    error('错误：未找到原始数据文件，请确认路径：%s', inputFile);
end
fprintf('========== 开始数据清洗 ==========\n');

% =============================================
% 1. 读取原始数据（强制全列读取为字符）
% =============================================
opts = detectImportOptions(inputFile, 'NumHeaderLines', 1);
opts = setvartype(opts, 'char'); 
T = readtable(inputFile, opts);

T.Properties.VariableNames = {
    'FullTitle', 'Department', 'Location', 'Category', 'Headcount', 'Date', 'Description'
};

% =============================================
% 2. 核心新增：自动修复数据列错位（逻辑顺序已完美修正）
% =============================================
fprintf('正在检查并修复列错位数据...\n');
fixedCount = 0;
for i = 1:height(T)
    hc = T.Headcount{i};
    cat = T.Category{i};
    loc = T.Location{i};
    dept = T.Department{i};
    
    % 如果 Headcount 变成了日期，说明右侧数据整体发生了左移错位
    if ~isempty(regexp(hc, '\d{4}-\d{2}-\d{2}', 'once'))
        fixedCount = fixedCount + 1;
        
        % 优先判断：如果原本应该是地点的 loc 列，变成了'技术'/'产品'，说明是缺失了【部门】
        if contains(loc, '技术') || contains(loc, '产品') || contains(loc, '设计') || contains(loc, '运营') || contains(loc, '销售')
            T.Date{i}       = hc;      % 日期归位
            T.Headcount{i}  = cat;     % 人数归位
            T.Category{i}   = loc;     % 类别归位 (技术)
            T.Location{i}   = dept;    % 地点归位 (北京市)
            T.Department{i} = '';      % 部门置空，后续第10步会自动推断填补
            
        % 否则就是缺失了【岗位类别】的情况
        else
            T.Date{i}       = hc;      % 日期归位
            T.Headcount{i}  = cat;     % 人数归位
            T.Category{i}   = '';      % 类别置空
            % Location 和 Department 本身没串列，保留原样即可
        end
    end
end
fprintf('已成功修复 %d 条错位数据。\n', fixedCount);

% =============================================
% 3. 删除完全重复的行与关键字段缺失的行
% =============================================
T = unique(T);
idxInvalid = cellfun(@(x) isempty(strtrim(x)), T.FullTitle) | cellfun(@(x) isempty(strtrim(x)), T.Description);
T(idxInvalid, :) = [];

% =============================================
% 4. 拆分岗位全称 -> 岗位名称 + 岗位编号
% =============================================
jobName = cell(height(T), 1);
jobID   = cell(height(T), 1);
for i = 1:height(T)
    title = T.FullTitle{i};
    tokens = regexp(title, '^(.*?)[（(](J\d+)[）)]$', 'tokens', 'once');
    if ~isempty(tokens)
        jobName{i} = strtrim(tokens{1});
        jobID{i}   = tokens{2};
    else
        jobName{i} = strtrim(title);
        jobID{i}   = '';
    end
end
T.JobName = jobName;
T.JobID   = jobID;

% =============================================
% 5. 清洗工作地点：提取主要城市
% =============================================
mainCity = cell(height(T), 1);
cityCount = zeros(height(T), 1);
for i = 1:height(T)
    loc = T.Location{i};
    if ~isempty(strtrim(loc))
        cityList = strsplit(loc, '，');
        if length(cityList) == 1
            cityList = strsplit(loc, ',');
        end
        mainCity{i} = strtrim(cityList{1});
        cityCount(i) = length(cityList);
    else
        mainCity{i} = '';
        cityCount(i) = 0;
    end
end
T.MainCity = mainCity;
T.CityCount = cityCount;

% =============================================
% 6. 清洗招聘人数："若干" -> NaN
% =============================================
headcountNum = NaN(height(T), 1);
for i = 1:height(T)
    hc = T.Headcount{i};
    if isempty(strtrim(hc)) || strcmp(strtrim(hc), '若干')
        continue;
    end
    numStr = regexp(hc, '\d+', 'match', 'once');
    if ~isempty(numStr)
        headcountNum(i) = str2double(numStr);
    end
end
T.HeadcountNum = headcountNum;

% =============================================
% 7. 清洗发布日期
% =============================================
parsedDate = NaT(height(T), 1);
for i = 1:height(T)
    dateStr = strtrim(T.Date{i});   
    if isempty(dateStr)
        continue;
    end
    try
        d = datetime(dateStr, 'InputFormat', 'yyyy-MM-dd HH:mm:ss', 'Locale', 'en_US');
    catch
        try d = datetime(dateStr, 'InputFormat', 'yyyy-MM-dd', 'Locale', 'en_US');
        catch
            try d = datetime(dateStr, 'InputFormat', 'yyyy/MM/dd', 'Locale', 'en_US');
            catch
                try d = datetime(dateStr, 'InputFormat', 'MM/dd/yyyy', 'Locale', 'en_US');
                catch d = datetime(dateStr, 'Locale', 'en_US');
                end
            end
        end
    end
    if ~isnat(d)
        parsedDate(i) = d;
    end
end
T.Date_Parsed = parsedDate; 
T.Year   = year(T.Date_Parsed);
T.Quarter= quarter(T.Date_Parsed);
T.Month  = month(T.Date_Parsed);

% =============================================
% 8. 清洗岗位描述
% =============================================
cleanedDesc = cell(height(T), 1);
for i = 1:height(T)
    txt = T.Description{i};
    if isempty(strtrim(txt))
        cleanedDesc{i} = '';
        continue;
    end
    txt = regexprep(txt, '<br\s*/?>', ' ');
    txt = regexprep(txt, '[\n\r\t]+', ' ');
    txt = regexprep(txt, '&[a-zA-Z]+;', ' ');
    txt = regexprep(txt, '[•·■◆\-]+', ' ');
    txt = regexprep(txt, '\t', ' ');
    txt = regexprep(txt, '\.{2,}', ' ');
    txt = regexprep(txt, '[、，,]', ' '); 
    txt = regexprep(txt, '\s+', ' ');
    cleanedDesc{i} = strtrim(txt);
end
T.CleanedDesc = cleanedDesc;

% =============================================
% 9. 提取关键技能
% =============================================
skillDict = {
    'Python',        'python|py|pandas|numpy|scikit|sklearn';
    'SQL',           'sql|hive|presto|spark sql|mysql|postgresql|结构化查询';
    'Java',          'java';
    'Cpp',           'c\+\+';  
    'Go',            'go|golang';
    'LLM',           '大模型|llm|大语言模型|生成式|文心|chatgpt|gpt|prompt|rag|agent|ai agent|智能体';
    'ML',            '机器学习|深度学习|ml|神经网络|cnn|rnn|transformer|强化学习|监督学习|无监督学习|人工智能';
    'DM',            '数据挖掘|dm|关联规则|聚类|分类|预测|异常检测|特征工程|挖掘';
    'DA',            '数据分析|ab测试|ab实验|归因分析|指标体系|漏斗分析|业务分析|经营分析';
    'BigData',       'hadoop|spark|flink|kafka|storm|dataflow|etl|数据管道|数据集成|大数据';
    'DB',            'mysql|redis|mongodb|postgresql|向量数据库|hbase|cassandra|elasticsearch|doris|clickhouse|数据库';
    'DataWarehouse', '数据仓库|dw|dwd|dws|数据湖|数据建模|数仓';
    'Cloud',         'docker|kubernetes|k8s|云原生|aws|azure|gcp|公有云|私有云|混合云|容器|云计算';
    'Frontend',      'react|vue|echarts|可视化|前端|html|css|javascript|web';
    'PM',            '产品设计|prd|需求分析|竞品分析|产品规划|产品经理|用户调研|产品运营';
};
nSkills = size(skillDict, 1);
skillNames = skillDict(:, 1);
skillPatterns = skillDict(:, 2);

for s = 1:nSkills
    colName = ['Skill_' skillNames{s}];
    T.(colName) = false(height(T), 1);
    pattern = skillPatterns{s};
    for i = 1:height(T)
        desc = T.CleanedDesc{i};
        if ~isempty(desc) && ~isempty(regexpi(desc, pattern, 'once'))
            T.(colName)(i) = true;
        end
    end
end
skillCols = T.Properties.VariableNames(startsWith(T.Properties.VariableNames, 'Skill_'));
T.SkillCount = sum(T{:, skillCols}, 2);

% =============================================
% 10. 提取一级事业群（BG）
% =============================================
bg = cell(height(T), 1);
for i = 1:height(T)
    dept = strtrim(T.Department{i});
    title = T.FullTitle{i};
    if isempty(dept) || strcmp(dept, '北京市')
        if contains(title, 'ACG') || contains(title, '智能云') || contains(title, '云')
            bg{i} = 'ACG(智能云)';
        elseif contains(title, 'IDG') || contains(title, '智能驾驶') || contains(title, '车')
            bg{i} = 'IDG(智能驾驶)';
        elseif contains(title, 'TPG') || contains(title, '技术中台')
            bg{i} = 'TPG(技术中台)';
        elseif contains(title, '小度')
            bg{i} = '小度科技';
        elseif contains(title, '安全') || contains(title, '效率平台')
            bg{i} = '安全与企业效率平台';
        else
            bg{i} = '其他';
        end
    elseif contains(dept, 'MEG')
        bg{i} = 'MEG(移动生态)';
    elseif contains(dept, 'ACG')
        bg{i} = 'ACG(智能云)';
    elseif contains(dept, 'IDG')
        bg{i} = 'IDG(智能驾驶)';
    elseif contains(dept, 'TPG')
        bg{i} = 'TPG(技术中台)';
    elseif contains(dept, '小度')
        bg{i} = '小度科技';
    elseif contains(dept, '安全与')
        bg{i} = '安全与企业效率平台';
    else
        bg{i} = dept;
    end
end
T.BG = bg;

% =============================================
% 11. 统计概览 & 整理最终表格
% =============================================
keepVars = {
    'FullTitle', 'Department', 'Location', 'Category', 'Headcount', 'Date', 'Description', ...
    'JobName', 'JobID', 'MainCity', 'CityCount', ...
    'HeadcountNum', 'Year', 'Quarter', 'Month', ...
    'CleanedDesc', ...
    skillCols{:}, ...
    'SkillCount', 'BG'
};
T_final = T(:, keepVars);

finalColNames = {
    '岗位全称', '部门', '工作地点', '岗位类别', '招聘人数', '发布日期', '岗位描述', ...
    '岗位名称', '岗位编号', '主要城市', '城市数量', ...
    '招聘人数_数值', '发布年份', '发布季度', '发布月份', ...
    '清洗后描述', ...
    skillCols{:}, ...
    '技能总数', '一级部门'
};
T_final.Properties.VariableNames = finalColNames;

% =============================================
% 12. 保存结果
% =============================================
writetable(T_final, outputExcel, 'WriteMode', 'overwritesheet');
writetable(T_final, outputCSV, 'Encoding', 'UTF-8');
fprintf('\n✅ 清洗后的数据已完美保存至：\n');
fprintf('   - %s\n', outputExcel);
disp('========== 处理完成 ==========');