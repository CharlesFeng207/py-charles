{
	"time_start":"2018/3/20", # 筛选开始时间
	"time_end":"2018/3/23", # 筛选截至时间
	"time_range_cols":"S,T", # 筛选目标列，有任意时间在区间内，则该行将被选中
	
	"src": "D:\Repositories\py_charles\examples\tracking\CD539 MCA PVPR Tracker (color impact) 20180308.xlsm", # 源表路径
	"src_sheet": "Tracking List", # 源表目标页
	"src_id_row": "3", # 源表那一列为数据id
	"src_data_start_row":"6", # 源表的数据是从哪一行开始的
	
	"working_table_name":"WeeklyWorkingStart", # 待生成的working表名字，留空则表示不需要生成该表
	"temp_working":"D:\Repositories\py_charles\examples\tracking\template_working.xlsx", # working模版表路径
	"temp_working_id_row":"1", # working模版表哪一行为数据id
	"temp_working_data_start_row":"2", # working模版表的数据是从哪一行开始的
	"working_combine_col":"E", # 生成working表时，会根据源表中该行的数据来合并数据行
	"working_combine_number_id":"Impact Parts No." # 合并后数量的数据id


	"delay_table_name":"WeeklyDelayStart", # 待生成的delay表名字，留空则表示不需要生成该表
	"temp_delay":"D:\Repositories\py_charles\examples\tracking\template_delay.xlsx", # delay模版表路径
	"temp_delay_id_row":"2", # delay模版表哪一行为数据id
	"temp_delay_data_start_row":"3", # delay模版表的数据是从哪一行开始的
	
	"delay_check_col":"U", # 生成delay表时，只会将源表中该行为空的数据行筛选出来
	"delay_combine_col":"E", # 生成delay表时，会根据源表中该行的数据来合并数据行
	"delay_combine_number_id":"Impact Parts No." # 合并后数量的数据id
}