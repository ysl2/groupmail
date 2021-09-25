# 群发邮件脚本

## 使用步骤

1. 在`main.py`文件中设置变量`FROM_ADDR`，`FROM_ALIAS`，`PASSWORD`。需要注意的是`PASSWORD`对于某些邮箱，是邮箱授权码(比如qq邮箱)。
2. 在`main.py`文件的第82行，设置邮箱的服务器。
3. 在`assets/list.xlsx`文件中设置需要群发的邮箱列表。`to`表示收件人，`cc`表示抄送人。一般情况下，直接按照样例文件设置即可(全部设置为`to`，不设置`cc`)。
4. 使用以下命令发送邮件。

	```python
	$ python main.py -s '题目202100925进展汇报' -c '邮件正文内容.txt' -a '附件20210925.pptx'
	```

## 参数列表

|参数|类型|含义|
|:---|---|---|
|`-s`, `--subject`|字符串|邮件主题|
|`-c`, `--content`|字符串(txt文件路径)|邮件正文(目前仅支持txt，即纯文本)|
|`-a`, `--attachment`|字符串(任意类型的文件路径)|邮件附件|