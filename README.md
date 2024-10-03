使用说明:
0. 确保config.json和Excel文件在程序同一目录下
1. 在config.json中设置Excel文件名和各项参数
2. 准备Excel文件,包含"联系人"和"发送内容"两个sheet
3. 在"联系人"sheet中填写联系人姓名和称呼
4. 在"发送内容"sheet的A1单元格填写要发送的消息内容
5. 运行程序时，选择是否要带称呼
6. 程序会自动读取Excel并控制微信发送消息
7. 发送过程中会在控制台显示进度和结果

注意事项:
- 使用前请确保微信PC客户端已登录并正常运行
- 联系人姓名必须与微信中的备注或昵称完全一致
- 发送过程中请勿手动操作微信,以免影响自动化流程
