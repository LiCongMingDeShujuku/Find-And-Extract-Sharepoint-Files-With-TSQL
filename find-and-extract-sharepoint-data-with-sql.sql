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
