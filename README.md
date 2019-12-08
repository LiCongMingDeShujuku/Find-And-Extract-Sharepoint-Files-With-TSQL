![CLEVER DATA GIT REPO](https://raw.githubusercontent.com/LiCongMingDeShujuku/git-resources/master/0-clever-data-github.png "李聪明的数据库")

# 使用TSQL搜索并提取Sharepoint文档
#### Find And Extract Sharepoint Files With TSQL

![#](images/Find-And-Extract-Sharepoint-Files-With-TSQL-01.png?raw=true "#")

## Contents

- [中文](#中文)
- [English](#English)
- [SQL Logic](#Logic)
- [Build Info](#Build-Info)
- [Author](#Author)
- [License](#License) 


## 中文
下面是一些可帮助你使用SQL查找提取，导出或以其他方式下载Sharepoint文件的快速SQL逻辑（logic）。 sp_oamethod已经存在了很长一段时间，并且它仍在工作。 你永远不知道MS什么时候会调用这些程序。

## English
Here is some quick SQL logic to help you find extract, export, or otherwise download Sharepoint documents with SQL. The sp_oamethod has been around for quite some time, and delighted it’s still working. You never know when MS will pull these procedures.

---
## Logic
```SQL
use [WSS_Content];
set nocount on
go
-- 	look for specific document files in sharepoint
-- 	查找SharePoint里的具体文件夹
select
    'database'  = db_name()
,   'time_created'  = left(alldocs.timecreated, 19)
,   'kb'        = (convert(bigint,alldocstreams.size))/1024
,   'mb'        = (convert(bigint,alldocstreams.size))/1024/1024
,   'list_name' = alllists.tp_title
,   'file_name' = alldocs.leafname
,   'url'       = alldocs.dirname
 
from
    alldocs join alldocstreams  on alldocs.id=alldocstreams.id 
    join alllists           on alllists.tp_id = alldocs.listid
where
    alldocstreams.[size] > 2048 -- looking for files under 2 mb
    and [alllists].[tp_title]   like '%training%'
    and [alldocs].[leafname]    like '%tutorial%'
    and right([alldocs].[leafname], 2) in ('oc', 'cx', 'df', 'sg', 'xt')
order by
    alldocs.timecreated desc
 
 
-- 	extract files using the leafname and dirname as identifiers
-- 	使用leafname和dirname标识符提取文档
declare @object_token       int
declare @destination_path   varchar(255)
declare @content_binary     varbinary(max)
set @destination_path   = '\\MyServerName\W$\Sharepoint_Extraction\my_tutorial.pdf'
 
select  @content_binary     = alldocstreams.content from alldocs join alldocstreams on alldocs.id = alldocstreams.id join alllists on alllists.tp_id = alldocs.listid
where  
    alldocs.leafname    = 'my_tutorial.pdf'
    and alldocs.dirname = 'Site/Process/Tutorials/TechnicalDocuments/How-To/Guides'
 
exec sp_oacreate 'adodb.stream', @object_token output
exec sp_oasetproperty   @object_token, 'type', 1
exec sp_oamethod    @object_token, 'open'
exec sp_oamethod    @object_token, 'write',     null, @content_binary
exec sp_oamethod    @object_token, 'savetofile',    null, @destination_path, 2
exec sp_oamethod    @object_token, 'close'
exec sp_oadestroy   @object_token
 
-- 	see files at the destination
-- 	最终看到查找的文档
declare @check_files    varchar(255)
set     @check_files    = '\\MyServerName\W$\Sharepoint_Extraction\'
exec    master..xp_dirtree @check_files, 1,1
go


```



[![WorksEveryTime](https://forthebadge.com/images/badges/60-percent-of-the-time-works-every-time.svg)](https://shitday.de/)

## Build-Info

| Build Quality | Build History |
|--|--|
|<table><tr><td>[![Build-Status](https://ci.appveyor.com/api/projects/status/pjxh5g91jpbh7t84?svg?style=flat-square)](#)</td></tr><tr><td>[![Coverage](https://coveralls.io/repos/github/tygerbytes/ResourceFitness/badge.svg?style=flat-square)](#)</td></tr><tr><td>[![Nuget](https://img.shields.io/nuget/v/TW.Resfit.Core.svg?style=flat-square)](#)</td></tr></table>|<table><tr><td>[![Build history](https://buildstats.info/appveyor/chart/tygerbytes/resourcefitness)](#)</td></tr></table>|

## Author

- **李聪明的数据库 Lee's Clever Data**
- **Mike的数据库宝典 Mikes Database Collection**
- **李聪明的数据库** "Lee Songming"

[![Gist](https://img.shields.io/badge/Gist-李聪明的数据库-<COLOR>.svg)](https://gist.github.com/congmingshuju)
[![Twitter](https://img.shields.io/badge/Twitter-mike的数据库宝典-<COLOR>.svg)](https://twitter.com/mikesdatawork?lang=en)
[![Wordpress](https://img.shields.io/badge/Wordpress-mike的数据库宝典-<COLOR>.svg)](https://mikesdatawork.wordpress.com/)

---
## License
[![LicenseCCSA](https://img.shields.io/badge/License-CreativeCommonsSA-<COLOR>.svg)](https://creativecommons.org/share-your-work/licensing-types-examples/)

![Lee Songming](https://raw.githubusercontent.com/LiCongMingDeShujuku/git-resources/master/1-clever-data-github.png "李聪明的数据库")

