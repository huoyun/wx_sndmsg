
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


config.jason内容如下：
{
    "excel_file_name": "微信批量发送消息配置表.xlsx",
    "wait_time_between_messages": 1.0,
    "wait_time_after_search": 0.5,
    "wait_time_after_enter": 0.5,
    "wait_time_for_chat_window": 1.0,
    "max_search_attempts": 2,
    "wait_time_between_attempts": 0.5,
    "wait_time_after_paste": 0.3,
    "wait_time_after_send": 0.3
}

配置参数说明：

excel_file_name: Excel文件的名称，包含联系人信息和消息内容。

wait_time_between_messages: 两条消息发送之间的等待时间（秒）。
    - 增加此值可以降低被检测为垃圾消息的风险，但会降低发送效率。

wait_time_after_search: 在搜索联系人后的等待时间（秒）。
    - 给微信一些时间来响应搜索操作。

wait_time_after_enter: 在按下回车键进入聊天窗口后的等待时间（秒）。
    - 确保聊天窗口已经打开。

wait_time_for_chat_window: 等待聊天窗口完全加载的时间（秒）。
    - 如果聊天记录较多，可能需要增加此值。

max_search_attempts: 尝试查找联系人的最大次数。
    - 增加此值可以提高找到联系人的成功率，但可能会降低效率。

wait_time_between_attempts: 两次查找尝试之间的等待时间（秒）。
    - 给微信一些时间来响应上一次的搜索操作。

wait_time_after_paste: 在粘贴消息内容后的等待时间（秒）。
    - 确保消息内容被正确粘贴。

wait_time_after_send: 在发送消息后的等待时间（秒）。
    - 给微信一些时间来发送消息。

注意：这些参数的最佳值可能因设备性能和网络条件而异。请根据实际使用情况进行调整。
