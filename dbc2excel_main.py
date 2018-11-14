#!/usr/bin/env python
import wx
import dbc2excel as d2e


class MyFrame(wx.Frame):
    """ We simply derive a new class of Frame. """
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(600, 500))
        self.path = ''
        #self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        #self.control = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        ##excel生成条件变量
        self.if_sig_desc = True
        self.if_sig_val_desc = True
        self.val_description_max_number = 70
        self.if_start_val = True
        self.if_recv_send = True
        self.if_asc_sort = True

        #静态文字
        self.quote = wx.StaticText(self, label="\n", pos=(420, 260))
        #控制台窗口
        self.logger = wx.TextCtrl(self, pos=(5, 300), size=(580, 130), style=wx.TE_MULTILINE | wx.TE_READONLY)

        # A button
        self.button =wx.Button(self, label="生成Excel文件", pos=(10, 150),size=(200,100))
        self.Bind(wx.EVT_BUTTON, self.create_excel,self.button)
        # B button
        b = wx.Button(self,-1,u"选择dbc文件",pos=(10, 20),size=(200,100))
        self.Bind(wx.EVT_BUTTON, self.select_file_button, b)
        # c button 增加图片
        #pic = wx.Image("./source/a.bmp", wx.BITMAP_TYPE_BMP).ConvertToBitmap()
        #c = wx.BitmapButton(self,-1,pic,pos=(250, 20),size=(290,150))
        #self.Bind(wx.EVT_BUTTON, self.select_file_button, b)

        #增加复选框
        #panel = wx.Panel(self)  # 创建画板，控件容器
        #信号描述
        HEIGHT = 25
        OFFSET = 20
        k = 1
        self.check1 = wx.CheckBox(self, -1, '生成信号描述', pos=(250, HEIGHT), size=(100, -1))
        self.Bind(wx.EVT_CHECKBOX, self.SigDescEvtCheckBox, self.check1)
        self.check1.Set3StateValue(True)

        self.check2 = wx.CheckBox(self, -1, '生成信号值描述', pos=(250, HEIGHT + k * OFFSET), size=(100, -1))
        self.Bind(wx.EVT_CHECKBOX, self.SigValDescEvtCheckBox, self.check2)
        self.check2.Set3StateValue(True)
        k += 1
        #最大信号值描述长度

        self.check3 = wx.CheckBox(self, -1, '生成初始值', pos=(250, HEIGHT + k * OFFSET), size=(100, -1))
        self.Bind(wx.EVT_CHECKBOX, self.StartValEvtCheckBox, self.check3)
        self.check3.Set3StateValue(True)
        k += 1
        self.check4 = wx.CheckBox(self, -1, '生成发送方和接收方', pos=(250, HEIGHT + k * OFFSET), size=(150, -1))
        self.Bind(wx.EVT_CHECKBOX, self.RecvSndEvtCheckBox, self.check4)
        self.check4.Set3StateValue(True)
        k += 1
        self.check5 = wx.CheckBox(self, -1, '升序排序(取消勾选降序)', pos=(250, HEIGHT + k * OFFSET), size=(150, -1))
        self.Bind(wx.EVT_CHECKBOX, self.SortEvtCheckBox, self.check5)
        self.check5.Set3StateValue(True)
        k += 1
        self.quote = wx.StaticText(self, label="信号值描述最大文本长度\n", pos=(250, HEIGHT + k * OFFSET), size = (140, 20))
        self.text1 = wx.TextCtrl(self, wx.ID_ANY, "70",pos=(400, HEIGHT + k * OFFSET), size=(100, 20), style=wx.TE_LEFT)
        #print(self.text1.Value)
        k += 1
        # Setting up the menu.
        filemenu = wx.Menu()
        # wx.ID_ABOUT and wx.ID_EXIT are standard IDs provided by wxWidgets.
        # wx.ID_ABOUT and wx.ID_EXIT are standard ids provided by wxWidgets.
        menuAbout = filemenu.Append(wx.ID_ABOUT, "&关于"," Information about this program")
        menuExit = filemenu.Append(wx.ID_EXIT,"E&xit"," Terminate the program")

        # Creating the menubar.
        menuBar = wx.MenuBar()
        menuBar.Append(filemenu,"&文件") # Adding the "filemenu" to the MenuBar
        self.SetMenuBar(menuBar)  # Adding the MenuBar to the Frame content.

        # Set events.
        self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)
        self.Bind(wx.EVT_MENU, self.OnExit, menuExit)
        self.logger.AppendText("Dbc转Excel工具仍在不断完善中\n***如果转换时间过长或无法转换，尝试取消勾选上方的选项再进行生成***\n有任何问题请在下面网址留言\nhttps://blog.csdn.net/hhlenergystory/article/details/80443454\n")

        self.Show(True)




    #响应事件
    def SigDescEvtCheckBox(self,event):
        self.if_sig_desc = not self.if_sig_desc
        #print(self.if_sig_desc)

    def SigValDescEvtCheckBox(self,event):
        self.if_sig_val_desc = not self.if_sig_val_desc
        #print(self.if_sig_val_desc)

    def StartValEvtCheckBox(self,event):
        self.if_start_val = not self.if_start_val
        #print(self.if_start_val)

    def RecvSndEvtCheckBox(self,event):
        self.if_recv_send = not self.if_recv_send
        #print(self.if_recv_send)
    def SortEvtCheckBox(self,event):
        self.if_asc_sort = not  self.if_asc_sort

    def OnAbout(self, e):
        # A message dialog box with an OK button. wx.OK is a standard ID in wxWidgets.
        dlg = wx.MessageDialog(self, "DBC转Excel工具\nBY黄洪磊 i2347\nV0.4\n有任何问题请发送dbc文件至int.honglei.huang@uaes.com", "关于", wx.OK)
        dlg.ShowModal()  # Show it
        dlg.Destroy()  # finally destroy it when finished.

    def OnExit(self, e):
        self.Close(True)  # Close the frame.

    def create_excel(self, event):
        self.logger.AppendText(" \n载入DBC文件完成\n" )
        dbc = d2e.DbcLoad(self.path)
        self.logger.AppendText(" 生成文件中！稍等... \n")
        if(str(self.text1.Value).isdigit()):
            self.val_description_max_number = int(self.text1.Value)
            #print(self.val_description_max_number)
        dbc.dbc2excel(self.path,self.if_sig_desc,self.if_sig_val_desc,self.val_description_max_number,self.if_start_val,self.if_recv_send,self.if_asc_sort)
        self.logger.AppendText(" 文件转换完成\n")

    def select_file_button(self, event):
        filesFilter = "Dicom (*.dbc)|*.dbc|" "All files (*.*)|*.*"
        fileDialog = wx.FileDialog(self, message="选择单个文件", wildcard=filesFilter, style=wx.FD_OPEN)
        dialogResult = fileDialog.ShowModal()
        if dialogResult != wx.ID_OK:
            return
        self.path = fileDialog.GetPath()
        self.logger.SetLabel('>>>选择文件：'+self.path)
    def OnEraseBack(self,event):
        dc = event.GetDC()
        if not dc:
            dc = wx.ClientDC(self)
            rect = self.GetUpdateRegion().GetBox()
            dc.SetClippingRect(rect)
        dc.Clear()
        bmp = wx.Bitmap("a.jpg")
        dc.DrawBitmap(bmp, 0, 0)

    #################
if __name__ == "__main__":
    app = wx.App(False)
    frame = MyFrame(None, 'DBC转Excel工具')
    app.MainLoop()
