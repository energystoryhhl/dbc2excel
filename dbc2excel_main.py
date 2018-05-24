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

        #静态文字
        self.quote = wx.StaticText(self, label="选择dbc文件 :", pos=(20, 30))
        #控制台窗口
        self.logger = wx.TextCtrl(self, pos=(5, 300), size=(580, 130), style=wx.TE_MULTILINE | wx.TE_READONLY)

        # A button
        self.button =wx.Button(self, label="生成Excel文件", pos=(100, 80))
        self.Bind(wx.EVT_BUTTON, self.create_excel,self.button)
        # B button
        b = wx.Button(self,-1,u"选择dbc文件",pos=(100, 20))
        self.Bind(wx.EVT_BUTTON, self.select_file_button, b)

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



        self.Show(True)
    #响应事件
    def OnAbout(self, e):
        # A message dialog box with an OK button. wx.OK is a standard ID in wxWidgets.
        dlg = wx.MessageDialog(self, "DBC转Excel工具\nBY黄洪磊 i2347\nV0.1", "关于", wx.OK)
        dlg.ShowModal()  # Show it
        dlg.Destroy()  # finally destroy it when finished.

    def OnExit(self, e):
        self.Close(True)  # Close the frame.

    def create_excel(self, event):
        self.logger.AppendText(" \nLoad DBC File \n" )
        dbc = d2e.DbcLoad(self.path)
        self.logger.AppendText(" 生成文件中！稍等... \n")
        dbc.dbc2excel(self.path)
        self.logger.AppendText(" 文件转换完成\n")

    def select_file_button(self, event):
        filesFilter = "Dicom (*.dbc)|*.dbc|" "All files (*.*)|*.*"
        fileDialog = wx.FileDialog(self, message="选择单个文件", wildcard=filesFilter, style=wx.FD_OPEN)
        dialogResult = fileDialog.ShowModal()
        if dialogResult != wx.ID_OK:
            return
        self.path = fileDialog.GetPath()
        self.logger.SetLabel('>>>选择文件：'+self.path)


    #################

app = wx.App(False)
frame = MyFrame(None, 'DBC转Excel工具')
app.MainLoop()