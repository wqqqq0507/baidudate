%% 招聘数据预处理Matlab 2024a专属修复版
% 清理工作区
clear; clc; close all;

%% ===================== 1. 读取原始数据（核心修复：强制保留原始列名） =====================
% 替换为你的Excel文件路径
file_path = '招聘数据_清洗后.xlsx';

% 核心修复：添加VariableNamingRule=preserve，强制保留Excel里的原始列名，不做任何修改
% 2024a版本完全支持该语法，彻底解决中文列名被修改的问题
raw_data = readtable(file_path, ...
    'TextType', 'string', ...
    'DatetimeType', 'datetime', ...
    'VariableNamingRule', 'preserve');

% 读取后先打印所有列名，确认和Excel里的列名完全一致
fprintf('===== 读取到的表格列名（和Excel完全一致）=====\n');
disp(raw_data.Properties.VariableNames');
fprintf('\n原始数据总行数: %d\n', height(raw_data));
fprintf('原始数据总列数: %d\n', width(raw_data));

% 复制数据到预处理表
df = raw_data;

%% ===================== 2. 处理缺失值（适配2024a版本+中文列名） =====================
% 2024a版本完全支持fillmissing，无需老版本兼容，直接用标准语法
% 部门列缺失值填充为'未知部门'
df.('部门') = fillmissing(df.('部门'), 'constant', '未知部门');
% 一级部门列缺失值填充为'未知一级部门'
df.('一级部门') = fillmissing(df.('一级部门'), 'constant', '未知一级部门');
% 招聘人数_数值缺失值填充为0
df.('招聘人数_数值') = fillmissing(df.('招聘人数_数值'), 'constant', 0);
% 城市数量缺失值填充为0
df.('城市数量') = fillmissing(df.('城市数量'), 'constant', 0);

% 数值列转换为整数类型
df.('招聘人数_数值') = int32(df.('招聘人数_数值'));
df.('城市数量') = int32(df.('城市数量'));

% 文本列缺失值填充为空字符串（完整覆盖所有文本列）
text_columns = [
    "岗位全称", "工作地点", "岗位类别", "招聘人数", "岗位描述", ...
    "岗位名称", "岗位编号", "主要城市", "清洗后描述"
];
for col = text_columns
    % 先判断列是否存在，避免列名不匹配报错
    if ismember(col, df.Properties.VariableNames)
        df.(col) = fillmissing(df.(col), 'constant', "");
    else
        fprintf('警告：列名【%s】在表格中未找到，跳过处理\n', col);
    end
end

%% ===================== 3. 处理异常值（适配招聘场景） =====================
% 招聘人数_数值异常值处理：大于100的设置为100
df.('招聘人数_数值')(df.('招聘人数_数值') > 100) = 100;
% 城市数量异常值处理：大于10的设置为10
df.('城市数量')(df.('城市数量') > 10) = 10;

%% ===================== 4. 数据类型转换（适配2024a版本） =====================
% 发布日期转换：匹配Excel里的日期格式，避免转换报错
df.('发布日期') = datetime(df.('发布日期'), 'InputFormat', 'yyyy-MM-dd HH:mm:ss', 'Errors', 'coerce');
% 发布年份、季度、月份转换为整数类型
date_columns = ["发布年份", "发布季度", "发布月份"];
for col = date_columns
    if ismember(col, df.Properties.VariableNames)
        df.(col) = fillmissing(df.(col), 'constant', 0);
        df.(col) = int32(df.(col));
    end
end

% 技能列转换为整数类型（0/1），完整覆盖所有Skill列
skill_columns = [
    "Skill_Python", "Skill_SQL", "Skill_Java", "Skill_Cpp", "Skill_Go", ...
    "Skill_LLM", "Skill_ML", "Skill_DM", "Skill_DA", "Skill_BigData", ...
    "Skill_DB", "Skill_DataWarehouse", "Skill_Cloud", "Skill_Frontend", ...
    "Skill_PM", "技能总数"
];
for col = skill_columns
    if ismember(col, df.Properties.VariableNames)
        df.(col) = int32(df.(col));
    end
end

%% ===================== 5. 文本数据清洗（健壮版） =====================
% 定义文本清洗函数，处理空值、多余空格、特殊字符
function clean_text = clean_text_func(text)
    if ismissing(text)
        clean_text = "";
        return;
    end
    % 去除多余空格（包括全角空格、换行符）
    clean_text = regexprep(text, '\s+', ' ');
    % 去除特殊字符，保留中文、英文、数字、常用标点
    clean_text = regexprep(clean_text, '[^\u4e00-\u9fa5a-zA-Z0-9，。！？；：、()（）\s]', '');
    % 去除首尾空格
    clean_text = strip(clean_text);
end

% 对文本列应用清洗函数
text_clean_columns = ["岗位描述", "清洗后描述", "岗位全称", "岗位名称"];
for col = text_clean_columns
    if ismember(col, df.Properties.VariableNames)
        df.(col) = arrayfun(@clean_text_func, df.(col));
    end
end

%% ===================== 6. 去重处理（按岗位编号去重） =====================
if ismember("岗位编号", df.Properties.VariableNames)
    [~, unique_idx] = unique(df.('岗位编号'), 'first');
    df = df(unique_idx, :);
    fprintf('\n去重后数据总行数: %d\n', height(df));
    fprintf('去重去除行数: %d\n', height(raw_data) - height(df));
else
    fprintf('\n警告：未找到【岗位编号】列，跳过去重处理\n');
end

%% ===================== 7. 新增特征列（和飞书表格逻辑一致） =====================
% 新增：是否为技术岗位
if ismember("岗位类别", df.Properties.VariableNames)
    df.('是否技术岗位') = int32(df.('岗位类别') == "技术");
end
% 新增：发布日期距离今天的天数
if ismember("发布日期", df.Properties.VariableNames)
    df.('发布天数') = int32(days(datetime('today') - df.('发布日期')));
    df.('发布天数') = fillmissing(df.('发布天数'), 'constant', 0);
end

%% ===================== 8. 结果保存与校验 =====================
% 保存预处理后的数据到Excel
output_path = '招聘数据_预处理完成_Matlab2024a修复版.xlsx';
writetable(df, output_path, 'SheetName', '预处理后数据', 'WriteVariableNames', true);

% 校验缺失值
missing_count = sum(ismissing(df), 'all');
fprintf('\n===== 预处理结果校验 =====\n');
fprintf('缺失值处理完成，所有列无缺失值: %s\n', missing_count == 0);
fprintf('预处理后文件已保存至: %s\n', output_path);

%% 可选：Excel样式美化（2024a版本完美支持）
% 打开Excel文件
excel_obj = actxserver('Excel.Application');
excel_obj.Visible = false;
workbook = excel_obj.Workbooks.Open(fullfile(pwd, output_path));
worksheet = workbook.Worksheets.Item('预处理后数据');

% 表头样式设置
header_range = worksheet.Range('1:1');
header_range.Font.Bold = true;
header_range.Font.Color = 16777215; % 白色
header_range.Interior.Color = 16737792; % 深蓝色
header_range.HorizontalAlignment = -4108; % 居中
header_range.RowHeight = 30;

% 数据区域样式设置
max_row = worksheet.UsedRange.Rows.Count;
max_col = worksheet.UsedRange.Columns.Count;

% 交替背景色
for row = 2:max_row
    if mod(row-2, 2) == 0
        bg_color = 15793151; % 浅蓝
    else
        bg_color = 16777215; % 白色
    end
    row_range = worksheet.Range(sprintf('%d:%d', row, row));
    row_range.Interior.Color = bg_color;
    row_range.Borders.Item(3).LineStyle = 1; % 底部细边框
    row_range.Borders.Item(3).Color = 11974323; % 浅灰
    row_range.RowHeight = 20;
end

% 列宽自动调整
worksheet.UsedRange.EntireColumn.AutoFit;

% 冻结窗格：冻结首行和前3列
worksheet.Range('D2').Select;
excel_obj.ActiveWindow.FreezePanes = true;

% 保存并关闭
workbook.Save;
workbook.Close;
excel_obj.Quit;
delete(excel_obj);

fprintf('Excel样式美化完成！\n');