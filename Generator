import wx
from wx.core import NO, VERTICAL, WXK_LEFT, Size
import pathlib, re,os
import pandas as pd

class BarcodePanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #Master Sizer for the Project
        self.MasterSizer =wx.BoxSizer(wx.VERTICAL)

        #Global Static Boxes Labels
        self.staticbox = wx.StaticBox(self,-1,"Barcode Generator Input", size=(408,75))
        self.staticboxsizer = wx.StaticBoxSizer(self.staticbox,wx.VERTICAL)
        self.box = wx.BoxSizer()
        
        #Item Static Box
        self.Itemstaticbox  = wx.StaticBox(self,-1,"Input Item Number")
        self.Itemstaticboxsizer = wx.StaticBoxSizer(self.Itemstaticbox,wx.VERTICAL)

        # User input for Item 
        self.Item = wx.TextCtrl(self, value="")
        self.Itemstaticboxsizer.Add(self.Item,1,wx.EXPAND|wx.ALL)
        self.box.Add(self.Itemstaticboxsizer,1,wx.EXPAND|wx.ALL)

        # A button
        self.button =wx.Button(self,-1 ,label="Generate and Save Barcode")
        self.Bind(wx.EVT_CHAR_HOOK, self.on_key)
        self.Bind(wx.EVT_BUTTON, self.OnClick,self.button)
        self.box.Add(self.button,1,wx.EXPAND|wx.ALL,10)

        #Adding Static boxes for input to Local Sizer
        self.staticboxsizer.Add(self.box,1,wx.EXPAND|wx.ALL,10)

        #Create Local Sizers for Results 
        self.Resultsstaticbox = wx.StaticBox(self,-1,"Generated Barcode",size=(408,75))
        self.Resultsstaticboxsizer = wx.StaticBoxSizer(self.Resultsstaticbox,wx.VERTICAL)
        self.Resultsbox = wx.BoxSizer()

        #create results Display
        self.logger = wx.TextCtrl(self,size=(350,75),style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.Resultsbox.Add(self.logger,wx.EXPAND|wx.ALL)
        self.Resultsstaticboxsizer.Add(self.Resultsbox,1,wx.EXPAND|wx.ALL,10)

        #Adding all UX Elements to sizers
        self.MasterSizer.Add(self.staticboxsizer)
        self.MasterSizer.Add(self.Resultsstaticboxsizer)
       
        #Display Information
        self.SetSizer(self.MasterSizer)
        self.Layout()

    def on_key(self, event):
        key = event.GetKeyCode()
        if key == wx.WXK_RETURN:
            self.button.SetFocus()
            self.button.SetDefault()
            self.button
        else:
            event.Skip()
        
    def OnClick(self,event):
        #Get Path to storage File
        if self.Item.Value == "":
            dial = wx.MessageDialog(None, 'Must input Item Name', 'Error', wx.OK)
            dial.ShowModal()
            return
        self.Path_Folder = pathlib.Path().resolve()
        self.Storage = "\Storage.txt"

        #Load Text Document and load data
        self.Data =pd.read_csv(str(self.Path_Folder)+self.Storage, sep='\t',header=0)
        self.Data['UPC Number'] = self.Data['UPC Number'].astype(str)
        self.UPC_List = list(self.Data['UPC Number'].unique())
        self.upc_list = []

        #remove all special characters
        for element in self.UPC_List:
            upc = re.sub(r'[^A-Za-z0-9]','',element)
            self.upc_list.append(upc)

        #sorting list of UPC's from lowest to highest
        self.upc_list.sort()
        self.last_upc = self.upc_list[-1]

        #Applying math for the "check Value"
        self.upc_11 = self.last_upc[0:11]
        self.upc_11_plus_one = str((int(self.upc_11)+1))
        self.upc_11_odd = self.upc_11_plus_one[0::2]
        self.upc_11_odd_list =[int(char) for char in self.upc_11_odd]
        self.odd_res = sum(self.upc_11_odd_list)*3       
        self.upc_11_even = self.upc_11_plus_one[1::2]
        self.upc_11_even_list = [int(char) for char in self.upc_11_even]
        self.even_res = sum(self.upc_11_even_list)
        self.odd_even_res = self.odd_res+self.even_res
        self.mod = self.odd_even_res % 10
        if self.mod == 0:
            self.check_value = 0
        else:
            self.check_value = 10 - self.mod
        self.generatedUpc = str(str(self.upc_11_plus_one) + str(self.check_value))

        if self.generatedUpc in self.upc_list:
            self.logger.AppendText("!!upc number already exists!!")
        self.logger.AppendText("{}{}{}{}".format(self.Item.Value,"\t",self.generatedUpc,"\n"))
        with open("Storage.txt", "a+") as file_object:
            file_object.seek(0)
            data= file_object.read(100)
            if len(data)>0:
                file_object.write("\n")
            file_object.write("{}\t{}".format(self.generatedUpc,self.Item.Value))    
            file_object.close()  
        self.Layout()

#-------------------------------------------Second Panel---------------------------------------------

class SecondPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent)
        #Master Sizer for the Project
        self.MasterSizer =wx.BoxSizer(wx.VERTICAL)

        #Global Boxsizers
        self.box = wx.BoxSizer()
        
        #Item Static Box
        self.Itemstaticbox  = wx.StaticBox(self,-1,"Select Excel File",size=(408,75))
        self.file_drop_target = FileDropTarget(self)
        self.FileDrop = wx.TextCtrl(self,value="Drag File Here",size=(90,20))
        self.FileDrop.SetDropTarget(self.file_drop_target)
        self.Itemstaticboxsizer = wx.StaticBoxSizer(self.Itemstaticbox,wx.VERTICAL)

        # User input for File 
        self.File = wx.DirPickerCtrl(self)
        self.Itemstaticboxsizer.Add(self.File,1,wx.EXPAND|wx.ALL)
        self.Itemstaticboxsizer.Add(self.FileDrop,0,wx.ALL|wx.CENTER)
        self.box.Add(self.Itemstaticboxsizer,1,wx.EXPAND|wx.ALL)

        # Load file button
        #Create Save Sizers
        self.Savestaticbox = wx.StaticBox(self,-1,"Save Barcodes",size=(408,75))
        self.Savestaticboxsizer = wx.StaticBoxSizer(self.Savestaticbox,wx.VERTICAL)
        self.Savebox = wx.BoxSizer()

        #Creating Save Display
        self.SaveButton= self.button = wx.Button(self,-1,label="Save Barcodes to Excel")
        self.Bind(wx.EVT_BUTTON, self.BuildList, self.button)
        self.SaveLocation = wx.DirPickerCtrl(self)
        self.Filename = wx.TextCtrl(self, value="Name of New File")

        #Adding Save Items to Display
        self.Savebox.Add(self.Filename,0,wx.EXPAND|wx.ALL,5)
        self.Savebox.Add(self.SaveLocation,1,wx.EXPAND|wx.ALL,5)
        self.Savestaticboxsizer.Add(self.Savebox,1,wx.EXPAND|wx.ALL,5)
        self.Savestaticboxsizer.Add(self.SaveButton,1,wx.EXPAND|wx.ALL,5)
    
        #Adding all UX Elements to sizers
        self.MasterSizer.Add(self.Itemstaticboxsizer)
        self.MasterSizer.Add(self.Savestaticboxsizer)

        #Display Information
        self.SetSizer(self.MasterSizer)
        self.Layout()
        
    def LoadStorageData(self):
        self.Path_Folder = pathlib.Path().resolve()
        self.Storage = "\Storage.txt"
        self.StorageData =pd.read_csv(str(self.Path_Folder)+self.Storage, sep='\t',header=0)
        self.StorageData['UPC Number'] = self.StorageData['UPC Number'].astype(str)
        self.StorageUPC = list(self.StorageData['UPC Number'].unique())
        self.StorageUPCList = []
        for element in self.StorageUPC:
            upc = re.sub(r'[^A-Za-z0-9]','',element)
            self.StorageUPCList.append(upc)
        self.StorageUPCList.sort()
        return self.StorageUPCList
    
    def LoadFileData(self):
        self.FilePath = self.File.GetPath()
        self.FileExtension = os.path.splitext(self.FilePath)
        self.FileType = self.FileExtension[1].lower()
        if self.FileType == ".csv":
            self.FileData =pd.read_csv(self.FilePath)
        else:
            self.FileData =pd.read_excel(self.FilePath)
        self.FileData['SKU'] = self.FileData['SKU'].astype(str)
        self.FileSKU = list(self.FileData['SKU'].unique())
        self.FileSKUList = []
        for element in self.FileSKU:
            sku = re.sub(r'[^A-Za-z0-9]','',element)
            self.FileSKUList.append(sku)
        return self.FileSKUList
        
    def GetLast(self,list):
        return list[-1]

    def GenNextUPC(self,upc):
        self.upc_11= upc[0:11]
        self.upc_11_plus_one = str((int(self.upc_11)+1))
        return self.upc_11_plus_one
        
    def GenCheckValue(self,upc):
        self.upc_11_odd = upc[0::2]
        self.upc_11_odd_list =[int(char) for char in self.upc_11_odd]
        self.odd_res = sum(self.upc_11_odd_list)*3
        self.upc_11_even = upc[1::2]
        self.upc_11_even_list = [int(char) for char in self.upc_11_even]
        self.even_res = sum(self.upc_11_even_list)
        self.odd_even_res = self.odd_res+self.even_res
        self.mod = self.odd_even_res % 10
        if self.mod == 0:
            self.check_value = 0
        else:
            self.check_value = 10 - self.mod
        return self.check_value
        
    def CompileUPC(self,upc11,checkValue):
        self.generatedUpc = str(str(upc11) + str(checkValue))
        return self.generatedUpc
        
    def WritetoFile(self,sku,upc):
        with open("Storage.txt", "a+") as file_object:
            file_object.seek(0)
            data= file_object.read(100)
            if len(data)>0:
                file_object.write("\n")
            file_object.write("{}\t{}".format(upc,sku))    
            file_object.close()
            
    def BuildList(self,event):
        self.savepath = self.SaveLocation.GetPath()
        FileName = str(self.Filename.Value())
        if self.savepath == '':
            dial = wx.MessageDialog(None, 'Please select where to save files', 'Info', wx.OK)
            dial.ShowModal()
            return
        elif FileName =="Name of New File" or FileName =="":
            dial = wx.MessageDialog(None, 'New file needs a name', 'Info', wx.OK)
            dial.ShowModal()
        ItemSet = []
        SkuList = self.LoadFileData()
        for sku in SkuList:
            UPClist = self.LoadStorageData()
            last = self.GetLast(UPClist)
            next = self.GenNextUPC(last)
            check = self.GenCheckValue(next)
            upc =self.CompileUPC(next,check)
            self.WritetoFile(sku,upc)
            Item = sku, upc
            ItemSet.append(Item)
        df_frame = self.BuildDataframe(ItemSet)
        self.Export(df_frame)       

    def BuildDataframe(self,Itemset):
        df = pd.DataFrame(Itemset)
        return df
        
    def Export(self,Dataframe):
        self.savepath = self.SaveLocation.GetPath()
        if self.savepath == '':
            dial = wx.MessageDialog(None, 'Please select where to save files', 'Info', wx.OK)
            dial.ShowModal()
            return
        FileName = str(self.Filename.Value())
        Dataframe.to_excel(self.savepath + FileName, sheet_name="Barcodes")
        dial = wx.MessageDialog(None, 'File Has been Saved', 'Info', wx.OK)
        dial.ShowModal()
       
    def SetInsertionPointEnd(self):
        self.FileDrop.SetInsertionPointEnd()
       
    def updateText(self, text):
            self.File.SetPath(text)

class FileDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window
    
    def OnDropFiles(self, x, y, filenames):
        self.window.SetInsertionPointEnd()
        for filepath in filenames:
            self.window.updateText(filepath)
        return True   

#needed to run in IDE
if __name__ == '__main__':
    app = wx.App(False)
    frame = wx.Frame(None, title="Generate Barcode Program",size=(430,300))
    nb = wx.Notebook(frame)
    nb.AddPage(BarcodePanel(nb), "Generate Barcode")
    nb.AddPage(SecondPanel(nb), "Import")
    frame.Show()
    app.MainLoop()
