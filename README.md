# pywinauto-for-Minitab18
使用pywinauto操作Minitab18进行CPK与GR&amp;R自动制作

需将Minitab18先打开,并打开Session的命令行(快捷键 Ctrl K)

运行py 输入CSV文件存放地址 点击OK获取文件

点击另一个按钮开始制作

制作时Minitab需在桌面最大化

跟据分辨率可能会出现GR&R制作时点击不到Session中图片的情况 此时需自行调整copy_image2_index

其他图表制作可以打开Session命令行 再手动制作一次获取出命令 然后进行cmd_focus.type_keys制作

