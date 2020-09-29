import os
import tkinter

from tkinter.filedialog import askopenfilename
import Modification
import sys

class Manager:
    def __init__(self):
        self.profile_path = os.environ["userprofile"]
        self.directory_path = None
        self.setting_file = "\\Settings.txt"
        self.plc_file = None
        self.graph_file = None
        self.modification_date = 0

    def is_setting_exist(self):
        if( os.path.isfile(self.profile_path + self.setting_file)):
            file = open(self.profile_path + self.setting_file, "r")
            content = file.readlines()
            file.close()
            if (len(content) == 3):
                return True
            os.remove(self.profile_path+self.setting_file)
        return False

    def file_dialog(self,plc_or_graph):
        if (plc_or_graph == 0):
            name = "PLC"
        else:
            name = "GRAPH"
        root = tkinter.Tk()
        root.withdraw()
        isFilled = False
        file_path = ""
        while not isFilled:
            file_path = askopenfilename(initialdir = self.profile_path,title = name,filetypes = (("CSV Files","*.csv"),("all files","*.*")))
            if (len(file_path) > 0):
                isFilled = True
        return file_path

    def get_files_path(self,filename):
        file = open(self.profile_path + self.setting_file,"r")
        content = file.readlines()
        file.close()
        if (filename == 1):
            return content[1].strip("\n")
        elif (filename==2):
            return content[2].strip("\n")
        else:
            return content[0].strip("\n")

    def save_settings(self,plc,graph,modification_time):
        file = open(self.profile_path + self.setting_file , "w")
        file.write(plc + "\n")
        file.write(graph + "\n")
        file.write(modification_time)
        file.close()


    def is_modified(self):
        new_modification_date = self.get_modification_time(self.plc_file)
        try:
            if (self.modification_date != new_modification_date):
                self.modification_date = new_modification_date
                self.save_settings(self.plc_file,self.graph_file,self.modification_date)
                return True
            return False
        except Exception:
            return False


    def set_modification_time(self):
        info = os.stat(self.plc_file)
        self.modification_date = str(info.st_mtime)

    def get_modification_time(self,plc_file):
        try:
            info = os.stat(plc_file)
        except Exception:
            print("File does-not exists")
            sys.exit(1)
        return str(info.st_mtime)


    def get_dir_path(self):
        return self.directory_path

    def set_settings(self):
        count =0
        if (self.is_setting_exist()):
            plc_file_path = self.get_files_path(0)
            graph_csv = self.get_files_path(1)
            modifiction_time = self.get_files_path(2)
            count = count +1
        else:
            plc_file_path = self.file_dialog(0)
            graph_csv = self.file_dialog(1)
            modifiction_time = self.get_modification_time(plc_file_path)
            self.save_settings(plc_file_path,graph_csv,modifiction_time)
        self.modification_date = modifiction_time
        self.plc_file = plc_file_path
        self.graph_file = graph_csv
        self.directory_path = os.path.dirname(plc_file_path)
        if(count == 0):
            return True
        elif (self.is_modified()):
            return True
        else:
            return False



    def run(self):
        settings = self.set_settings()
        if(settings):
            create_files = Modification.Modifcation(self.plc_file,self.graph_file)
            create_files.run()
        else:
            print("The file has not changed")




if __name__ == "__main__":
    x = Manager()
    x.run()
    print("yes")


