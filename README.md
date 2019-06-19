# lte-parameter-autocheck
LTE参数自动检查

1、备份：
运行Compression.exe。
源数据放在“template”；
压缩文件放在“压缩文件”；（不能删，很重要！）
合并后的最终文件放在“合并文件”——最后请把全网的表格命名为“全网参数备份20XXXXXX（当天的日期）”。

2、检查：
运行Inspection.exe(由py2exe生成，实际是__main__.py)。
结果放在当前文件夹，命名为“LTE参数检查结果.csv”。

3、注意：（这是最重要的！）
因为本程序使用的类库是只支持excel2003格式的，所以存在一定限制，在使用时请注意以下要点：
——先在网管导出是excel2007格式的表格，那么在做完第一步“备份”之后（一定要等合并完再做后续的操作），必须先把“压缩文件”里的所有excel2007格式的文件转为excel2003格式（原excel2007格式文件可保留），再运行Inspection.exe。（推荐） 
~   

$\bigcup((\forall x \in \{全网/分省业务, x \ge s*1% \}) \bigcap (\forall y \in \{全网/分省业务, y \neq x & y \ge s*1% \})) / 本省通信客户数(s)$
