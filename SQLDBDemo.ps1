#配置信息
$Database   = 'EOL_Generic'
$Server     = '"CNSZPNB712\SQLEXPRESS"'
$UserName   = 'sa'
$Password   = '1qaz@WSX#EDC'
 
#创建连接对象
$SqlConn = New-Object System.Data.SqlClient.SqlConnection
 
#使用账号连接MSSQL
$SqlConn.ConnectionString = "Data Source=$Server;Initial Catalog=$Database;user id=$UserName;pwd=$Password"
 
#或者以 windows 认证连接 MSSQL
#$SqlConn.ConnectionString = "Data Source=$Server;Initial Catalog=$Database;Integrated Security=SSPI;"
 
#打开数据库连接
$SqlConn.open()
 
#执行语句方法一
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.connection = $SqlConn
$SqlCmd.commandtext = 'delete top(1) from dbo.EOL'
$SqlCmd.executenonquery()

<# 设置备份权限
USE [EOL_Generic]
GO
ALTER ROLE [db_backupoperator] ADD MEMBER [BUILTIN\Users]
GO
#>
 
#执行语句方法二
#$SqlCmd = $SqlConn.CreateCommand()
#$SqlCmd.commandtext = 'delete top(1) from dbo.B'
#$SqlCmd.ExecuteScalar()
 
#方法三，查询显示
#$SqlCmd.commandtext = 'select name,recovery_model_desc,log_reuse_wait_desc from sys.databases'
#$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
#$SqlAdapter.SelectCommand = $SqlCmd
#$set = New-Object data.dataset
#$SqlAdapter.Fill($set)
#$set.Tables[0] | Format-Table -Auto
 
#关闭数据库连接
$SqlConn.close()